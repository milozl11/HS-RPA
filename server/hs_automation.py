"""SAP GTS Fiori automation for "Classify Products" (HS/tariff maintenance).

Flow driven here:
  1. Selection screen -> Get Variant... -> clear Created By, set variant name
  2. Execute the variant -> worklist
  3. Export the worklist to spreadsheet (download captured to disk)
  4. For each approved HS group: paste its products into the Product
     "Multiple Selection" popup (Shift+F12 from the OS clipboard, confirmed
     with F8), execute, select all rows, then fill the tariff code

The final commit (Start Mass Classification / F8) is refused unless
`allow_commit` is explicitly true. The default is a dry run that stops with the
dialog filled, which is what production validation requires.
"""

from __future__ import annotations

import ctypes
import queue
import re
import threading
import time
import traceback
from collections.abc import Callable
from pathlib import Path
from typing import Any
from urllib.parse import urlsplit

from hs_reference import HsReference
from hs_worklist import REASON_LABELS, analyze, build_description_plan, read_worklist
from playwright.sync_api import Frame, Page, sync_playwright
from playwright.sync_api import TimeoutError as PWTimeout
from sap_automation import (
    BrowserBusyError,
    _css_escape,
    _entries_by_country,
    _origin_of,
    _try_form_login,
    cleanup_profile_lock,
)
from sap_automation import (
    _mass_maintain_selection as _mass_maintain_description_selection,
)
from sap_automation import (
    _process_one_material as _process_description_material,
)
from sap_automation import (
    _set_display_maintained as _set_description_display_maintained,
)
from sap_automation import (
    _wait_for_worklist_or_empty as _wait_for_description_worklist,
)

_LOCK_FILES = ("SingletonLock", "SingletonCookie", "SingletonSocket")

_NO_SELECTION_RE = re.compile(
    r"no lines were selected|keine zeilen|no data selected", re.IGNORECASE
)
_PRODUCT_CLASSIFIED_RE = re.compile(
    r"Product\s+([0-9A-Za-z._/-]+)\s+(?:is\s+)?classified\b",
    re.IGNORECASE,
)


class HsSafetyError(RuntimeError):
    """Stop the remaining groups because selection integrity is uncertain."""


class HsEmptyWorklistError(RuntimeError):
    """The variant returned no rows; the scheme has nothing to maintain."""


class HsMissingVariantError(RuntimeError):
    """SAP has no matching Get Variant result for this scheme."""


def commercial_description_url(classification_url: str) -> str:
    """Use the same SAP system/client for the commercial-description app."""
    base = str(classification_url or "").split("#", 1)[0]
    if not base:
        raise ValueError("URL-ul SAP pentru clasificare lipseste.")
    return base + "#CustomsProduct-manageCustomsDescription?sap-ui-tech-hint=GUI"


def _root() -> Path:
    return Path(__file__).resolve().parent.parent


def _user_data_dir() -> Path:
    # Separate Chromium profile so the descriptions app and this app can run
    # independently without fighting over the same Singleton lock.
    return _root() / "user-data-hs"


def _downloads_dir() -> Path:
    return _root() / "downloads"


def _user_data_locked() -> bool:
    d = _user_data_dir()
    return any((d / f).exists() for f in _LOCK_FILES)


def cleanup_hs_lock(kill_chrome: bool = True) -> dict:
    return cleanup_profile_lock(_user_data_dir(), kill_chrome)


def _set_windows_clipboard_text(text: str) -> None:
    """Write UTF-16 text to the Windows clipboard for SAP Shift+F12 import."""
    user32 = ctypes.WinDLL("user32", use_last_error=True)
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

    open_clipboard = user32.OpenClipboard
    open_clipboard.argtypes = [ctypes.c_void_p]
    open_clipboard.restype = ctypes.c_int

    close_clipboard = user32.CloseClipboard
    close_clipboard.argtypes = []
    close_clipboard.restype = ctypes.c_int

    empty_clipboard = user32.EmptyClipboard
    empty_clipboard.argtypes = []
    empty_clipboard.restype = ctypes.c_int

    set_clipboard_data = user32.SetClipboardData
    set_clipboard_data.argtypes = [ctypes.c_uint, ctypes.c_void_p]
    set_clipboard_data.restype = ctypes.c_void_p

    global_alloc = kernel32.GlobalAlloc
    global_alloc.argtypes = [ctypes.c_uint, ctypes.c_size_t]
    global_alloc.restype = ctypes.c_void_p

    global_lock = kernel32.GlobalLock
    global_lock.argtypes = [ctypes.c_void_p]
    global_lock.restype = ctypes.c_void_p

    global_unlock = kernel32.GlobalUnlock
    global_unlock.argtypes = [ctypes.c_void_p]
    global_unlock.restype = ctypes.c_int

    global_free = kernel32.GlobalFree
    global_free.argtypes = [ctypes.c_void_p]
    global_free.restype = ctypes.c_void_p

    cf_unicode_text = 13
    gmem_moveable = 0x0002
    normalized = text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
    payload = normalized + "\0"
    encoded = payload.encode("utf-16-le")

    last_error = 0
    for _ in range(10):
        if open_clipboard(None):
            handle = None
            try:
                if not empty_clipboard():
                    last_error = ctypes.get_last_error()
                    raise OSError(last_error, "EmptyClipboard failed")
                handle = global_alloc(gmem_moveable, len(encoded))
                if not handle:
                    last_error = ctypes.get_last_error()
                    raise OSError(last_error, "GlobalAlloc failed")
                locked = global_lock(handle)
                if not locked:
                    last_error = ctypes.get_last_error()
                    raise OSError(last_error, "GlobalLock failed")
                try:
                    ctypes.memmove(locked, encoded, len(encoded))
                finally:
                    global_unlock(handle)

                if not set_clipboard_data(cf_unicode_text, handle):
                    last_error = ctypes.get_last_error()
                    raise OSError(last_error, "SetClipboardData failed")
                handle = None  # ownership transferred to the OS clipboard
                return
            finally:
                if handle:
                    global_free(handle)
                close_clipboard()
        last_error = ctypes.get_last_error()
        time.sleep(0.12)

    raise RuntimeError(
        f"Nu pot accesa clipboard-ul Windows pentru importul SAP (winerr={last_error})."
    )


# ============================================================================
# SAP screen helpers
# ============================================================================


def _classify_frame(page: Page, timeout_s: float = 30.0) -> Frame:
    """Frame holding the Classify Products GUI (selection screen or worklist)."""
    end = time.time() + timeout_s
    last_err = ""
    while time.time() < end:
        best_frame = None
        best_score = 0
        for f in page.frames:
            try:
                probe = _frame_probe(f)
                if _frame_is_ready(probe) and probe.get("score", 0) > best_score:
                    best_frame = f
                    best_score = probe["score"]
            except Exception as e:
                last_err = str(e)[:120]
                continue
        if best_frame is not None:
            return best_frame
        page.wait_for_timeout(400)
    raise TimeoutError(
        f"Frame-ul SAP 'Classify Products' nu a fost gasit. URL: {page.url} "
        f"(last_err={last_err})"
    )


def _frame_is_ready(probe: dict) -> bool:
    selection = bool(probe.get("hasVariant") and probe.get("hasScheme"))
    worklist = bool(
        probe.get("hasWorklist")
        and (probe.get("hasClassify") or probe.get("hasExec") or probe.get("known"))
    )
    return selection or worklist


def _frame_probe(frame: Frame) -> dict:
    """Return non-sensitive markers used to diagnose SAP frame selection."""
    return frame.evaluate(
        """() => {
            const visible = el => !!el && el.offsetParent !== null;
            const text = (document.body && document.body.innerText) || '';
            const hasVariant = Array.from(document.querySelectorAll(
                '[title*="Get Variant" i]'
            )).some(visible);
            const hasClassify = Array.from(document.querySelectorAll(
                '[title*="Classify Products" i], [title*="Start Mass Classification" i]'
            )).some(visible);
            const inputLabel = input => {
                const parts = [
                    input.getAttribute('title') || '',
                    input.getAttribute('aria-label') || '',
                ];
                for (const id of (input.getAttribute('aria-labelledby') || '')
                    .split(/\\s+/).filter(Boolean)) {
                    const label = document.getElementById(id);
                    if (label) parts.push(label.textContent || '');
                }
                if (input.id) {
                    const label = document.querySelector(
                        'label[for="' + CSS.escape(input.id) + '"]'
                    );
                    if (label) parts.push(label.textContent || '');
                }
                return parts.join(' ');
            };
            const hasScheme = Array.from(document.querySelectorAll('input')).some(
                input => visible(input) && /numbering\\s*scheme|nummernschema/i.test(
                    inputLabel(input)
                )
            );
            const hasWorklist = Array.from(document.querySelectorAll(
                'table, [role="grid"]'
            )).some(grid => visible(grid) && /\\bProduct\\b/i.test(
                grid.textContent || ''
            ));
            const hasExec = Array.from(document.querySelectorAll(
                '[title*="Execute" i]'
            )).some(visible);
            const known = /Save as Variant|Numbering Scheme|Classify Products/i.test(text);
            return {
                hasVariant,
                hasScheme,
                hasClassify,
                hasWorklist,
                hasExec,
                known,
                score: (hasVariant ? 8 : 0) +
                    (hasScheme ? 5 : 0) +
                    (hasClassify ? 6 : 0) +
                    (hasWorklist ? 4 : 0) +
                    (hasExec && known ? 2 : 0),
            };
        }"""
    )


def _safe_page_url(url: str) -> str:
    """Keep only host/path/hash intent; never log query or credentials."""
    try:
        parsed = urlsplit(str(url or ""))
        host = parsed.hostname or ""
        path = parsed.path or "/"
        fragment = parsed.fragment.split("?", 1)[0][:80]
        shown = f"{parsed.scheme}://{host}{path}" if parsed.scheme else path
        if fragment:
            shown += f"#{fragment}"
        return shown
    except Exception:
        return "(url indisponibil)"


def _context_pages(page: Page) -> list[Page]:
    """The launchpad can open the WebGUI transaction in a second browser tab."""
    pages = [page]
    try:
        for other in page.context.pages:
            if other is not page and not other.is_closed():
                pages.append(other)
    except Exception:
        pass
    return pages


def _find_frame_any_page(
    page: Page,
    frame_finder: Callable[..., Frame],
    timeout_s: float,
) -> Frame:
    last_exc: Exception | None = None
    for candidate in _context_pages(page):
        try:
            return frame_finder(candidate, timeout_s=timeout_s)
        except Exception as exc:
            last_exc = exc
    raise last_exc or TimeoutError("Nicio pagina SAP disponibila.")


def _frames_debug(page: Page) -> str:
    """Frame count plus paths, so an empty WebGUI shell is visible in logs."""
    try:
        paths = []
        for frame in page.frames[:6]:
            try:
                paths.append((urlsplit(frame.url).path or "/")[-40:])
            except Exception:
                paths.append("?")
        return f"{len(page.frames)} [{', '.join(paths)}]"
    except Exception:
        return "?"


def _frame_debug_label(page: Page, frame: Frame) -> str:
    try:
        index = page.frames.index(frame)
    except ValueError:
        index = -1
    try:
        parsed = urlsplit(frame.url)
        location = parsed.path or "/"
    except Exception:
        location = "?"
    return f"index={index}, path={location[:120]}"


_VARIANT_DIALOG_SELECTOR = (
    '[role=dialog], [id^="webguiPopupWindow"], .lsPWNew, .LS_PopupWindow2'
)


def _wait_until_idle(page: Page, timeout_s: float = 6.0) -> bool:
    """Wait for SAP to stop showing a busy overlay instead of sleeping blindly."""
    deadline = time.time() + timeout_s
    while time.time() < deadline:
        busy = False
        for frame in page.frames:
            try:
                busy = busy or bool(
                    frame.evaluate(
                        """() => Array.from(document.querySelectorAll(
                            '[aria-busy="true"], .lsBusy, .lsProgress, .sapUiLocalBusyIndicator'
                        )).some(el => el && el.offsetParent !== null)"""
                    )
                )
            except Exception:
                continue
        if not busy:
            return True
        page.wait_for_timeout(150)
    return False


def _find_variant_dialog(
    page: Page,
    preferred: Frame,
    timeout_s: float = 15.0,
) -> tuple[Frame, Any]:
    """Find the visible Get Variant dialog, including parent/child frames."""
    selector = _VARIANT_DIALOG_SELECTOR
    deadline = time.time() + timeout_s
    while time.time() < deadline:
        frames = [preferred] + [frame for frame in page.frames if frame != preferred]
        for frame in frames:
            try:
                marked = frame.evaluate(
                    """(selector) => {
                        const visible = el => !!el && el.offsetParent !== null;
                        document.querySelectorAll('[data-rpa-variant-dialog]').forEach(
                            el => el.removeAttribute('data-rpa-variant-dialog')
                        );
                        const dialogs = Array.from(document.querySelectorAll(selector))
                            .filter(visible)
                            .filter(dialog => {
                                const text = (dialog.textContent || '').replace(/\\s+/g, ' ');
                                return dialog.querySelector(
                                    'input[title="Variant Name"], input[aria-label="Variant Name"]'
                                ) || /get variant|variant name|created by/i.test(text);
                            });
                        const leaves = dialogs.filter(dialog => !dialogs.some(
                            other => other !== dialog && other.contains(dialog)
                        ));
                        const hit = leaves[leaves.length - 1];
                        if (!hit) return { found: false, visible: dialogs.length };
                        hit.setAttribute('data-rpa-variant-dialog', '1');
                        return {
                            found: true,
                            id: hit.id || null,
                            sample: (hit.textContent || '').replace(/\\s+/g, ' ').trim().slice(0, 100),
                        };
                    }""",
                    selector,
                )
                if marked.get("found"):
                    return frame, frame.locator('[data-rpa-variant-dialog="1"]').first
            except Exception:
                continue
        page.wait_for_timeout(250)
    raise TimeoutError(
        "Dialogul Get Variant nu a devenit vizibil in niciun frame SAP dupa click."
    )


def _commercial_description_frame(page: Page, timeout_s: float = 30.0) -> Frame:
    """Find the description app, rejecting a stale Classify Products frame.

    Both selection screens expose Product and a numbering-scheme field, so the
    'Display Maintained Products' checkbox is the marker that identifies the
    description screen, and 'Start Mass Classification' the classifier.
    """
    end = time.time() + timeout_s
    last_err = ""
    last_probe = ""
    while time.time() < end:
        for frame in page.frames:
            try:
                probe = frame.evaluate(
                    """() => {
                        const visible = el => !!el && el.offsetParent !== null;
                        const text = (document.body && document.body.innerText) || '';
                        const labeled = (el) => (
                            (el.getAttribute('title') || '') + ' ' +
                            (el.getAttribute('aria-label') || '') + ' ' +
                            (el.textContent || '')
                        );
                        const product = Array.from(document.querySelectorAll(
                            'input[title="Product"], input[aria-label="Product"], ' +
                            'input[title="Produkt"], input[aria-label="Produkt"]'
                        )).some(visible);
                        const displayMaintained = Array.from(
                            document.querySelectorAll(
                                '[role="checkbox"], input[type="checkbox"]'
                            )
                        ).some(el => visible(el) &&
                            /Display Maintained|gepflegte Produkte/i.test(labeled(el)));
                        const descriptionTitle =
                            /Customs Commercial Description|Kommerzielle Beschreibung/i.test(text);
                        const massMaintenance = Array.from(
                            document.querySelectorAll(
                                '[title*="Start Mass Maintenance" i], ' +
                                '[aria-label*="Start Mass Maintenance" i]'
                            )
                        ).some(visible);
                        const multipleSelection = Array.from(
                            document.querySelectorAll(
                                '[title="Multiple Selection"], [aria-label="Multiple Selection"]'
                            )
                        ).some(visible);
                        const execute = Array.from(document.querySelectorAll(
                            '[title*="Execute" i], [aria-label*="Execute" i], ' +
                            '[title*="Ausfuehren" i], [title*="AusfÃ¼hren" i]'
                        )).some(visible);
                        const worklistField = Array.from(document.querySelectorAll(
                            'input[title*="Worklist" i], input[aria-label*="Worklist" i], ' +
                            'input[title*="Product Master Worklist" i]'
                        )).some(visible);
                        const logicalSystemGroup = Array.from(
                            document.querySelectorAll(
                                'input[title*="Logical System Group" i], ' +
                                'input[aria-label*="Logical System Group" i]'
                            )
                        ).some(visible);
                        const schemeField = Array.from(document.querySelectorAll(
                            'input[title*="Numbering Scheme" i], ' +
                            'input[aria-label*="Numbering Scheme" i], ' +
                            'input[title*="Nummernschema" i]'
                        )).some(visible);
                        const massClassification = Array.from(
                            document.querySelectorAll(
                                '[title*="Start Mass Classification" i], ' +
                                '[aria-label*="Start Mass Classification" i]'
                            )
                        ).some(visible);
                        const classifier = massClassification;
                        // Only the description selection screen offers
                        // 'Display Maintained Products'.
                        const descriptionScreen = displayMaintained ||
                            descriptionTitle || massMaintenance;
                        return {
                            ready: product && !classifier && descriptionScreen,
                            product,
                            classifier,
                            displayMaintained,
                            descriptionTitle,
                            massMaintenance,
                            multipleSelection,
                            execute,
                            worklistField,
                            logicalSystemGroup,
                            schemeField,
                            massClassification,
                        };
                    }"""
                )
                last_probe = (
                    f"product={int(bool(probe.get('product')))},"
                    f"disp={int(bool(probe.get('displayMaintained')))},"
                    f"worklist={int(bool(probe.get('worklistField')))},"
                    f"lsg={int(bool(probe.get('logicalSystemGroup')))},"
                    f"scheme={int(bool(probe.get('schemeField')))},"
                    f"massClass={int(bool(probe.get('massClassification')))},"
                    f"exec={int(bool(probe.get('execute')))},"
                    f"multi={int(bool(probe.get('multipleSelection')))}"
                )
                if probe.get("ready"):
                    return frame
            except Exception as exc:
                last_err = str(exc)[:120]
        page.wait_for_timeout(300)
    raise TimeoutError(
        "Frame-ul SAP 'Manage Customs Commercial Descriptions' nu a fost "
        f"gasit. URL: {_safe_page_url(page.url)} "
        f"(probe={last_probe or '-'} last_err={last_err})"
    )


def _click_by_title(gf: Frame, pattern: str, timeout_ms: int = 8000) -> None:
    gf.locator(f'[title*="{pattern}" i]').first.click(timeout=timeout_ms)


_NO_VARIANT_RE = re.compile(
    r"no variants found(?: for this selection)?|keine variante(?:n)? gefunden",
    re.IGNORECASE,
)


def _visible_sap_dialog_texts(page: Page, limit: int = 4) -> list[str]:
    texts: list[str] = []
    for frame in page.frames:
        try:
            found = frame.evaluate(
                """() => Array.from(document.querySelectorAll(
                    '[role=dialog], [id^="webguiPopupWindow"], .lsPWNew, .LS_PopupWindow2'
                )).filter(d => d.offsetParent).map(d =>
                    (d.textContent || '').replace(/\\s+/g, ' ').trim().slice(0, 180)
                )"""
            )
        except Exception:
            continue
        for item in found or []:
            if item and item not in texts:
                texts.append(str(item))
            if len(texts) >= limit:
                return texts
    return texts


def _dialog_texts_indicate_missing_variant(texts) -> bool:
    return any(_NO_VARIANT_RE.search(str(text or "")) for text in (texts or []))


def _dismiss_variant_search_dialogs(
    page: Page, log: Callable[[str, str], None]
) -> None:
    """Close Information + Get Variant so the next scheme can start cleanly."""
    for _ in range(4):
        if not _visible_sap_dialog_texts(page):
            return
        clicked = False
        for frame in page.frames:
            try:
                clicked = bool(
                    frame.evaluate(
                        """() => {
                            const visible = el => el && el.offsetParent !== null;
                            const dialogs = Array.from(document.querySelectorAll(
                                '[role=dialog], [id^="webguiPopupWindow"], .lsPWNew, .LS_PopupWindow2'
                            )).filter(visible);
                            for (const dialog of dialogs) {
                                const buttons = Array.from(dialog.querySelectorAll(
                                    '[title], [aria-label], button, [role=button]'
                                )).filter(visible);
                                for (const btn of buttons) {
                                    const label = (
                                        (btn.getAttribute('title') || '') + ' ' +
                                        (btn.getAttribute('aria-label') || '') + ' ' +
                                        (btn.textContent || '')
                                    ).replace(/\\s+/g, ' ').trim();
                                    if (/continue|close|cancel|ok|^x$/i.test(label)) {
                                        btn.click();
                                        return true;
                                    }
                                }
                            }
                            return false;
                        }"""
                    )
                )
            except Exception:
                continue
            if clicked:
                break
        if not clicked:
            page.keyboard.press("Enter")
        page.wait_for_timeout(250)
        if _visible_sap_dialog_texts(page):
            page.keyboard.press("Escape")
            page.wait_for_timeout(250)
    leftover = _visible_sap_dialog_texts(page)
    if leftover:
        log("warn", f"Dialog SAP inca vizibil dupa varianta lipsa: {leftover[:1]}")


def _raise_if_missing_variant(
    page: Page,
    variant: str,
    remaining: list[str],
    log: Callable[[str, str], None],
) -> None:
    if not _dialog_texts_indicate_missing_variant(remaining):
        return
    log(
        "warn",
        f"SAP nu are varianta {variant!r}; inchid dialogul si trec la schema urmatoare.",
    )
    _dismiss_variant_search_dialogs(page, log)
    raise HsMissingVariantError(f"Nicio varianta SAP gasita pentru {variant!r}.")


def _apply_variant(
    page: Page,
    gf: Frame,
    variant: str,
    log: Callable[[str, str], None],
) -> Frame:
    """Get Variant... -> clear Created By -> set variant name -> Execute."""
    log("info", f"Deschid 'Get Variant...' pentru varianta {variant!r}")
    _click_by_title(gf, "Get Variant")
    log("info", "Click Get Variant trimis; astept dialogul SAP")
    dialog_frame, dialog = _find_variant_dialog(page, gf, timeout_s=15)
    log(
        "info",
        f"Dialog Get Variant vizibil ({_frame_debug_label(page, dialog_frame)})",
    )

    # The variant search dialog defaults 'Created By' to the current user, which
    # hides shared standard variants. It must be cleared.
    cleared = dialog_frame.evaluate(
        """() => {
            const user = document.querySelector('input[title="User Name"]');
            if (!user) return false;
            user.focus();
            user.value = '';
            user.setAttribute('value', '');
            user.dispatchEvent(new Event('input', { bubbles: true }));
            user.dispatchEvent(new Event('change', { bubbles: true }));
            return true;
        }"""
    )
    log("info", f"Camp 'Created by' golit: {'DA' if cleared else 'NU'}")

    name_box = dialog_frame.locator(
        'input[title="Variant Name"], input[aria-label="Variant Name"]'
    ).first
    name_box.wait_for(state="visible", timeout=8000)
    name_box.click()
    name_box.fill("")
    page.keyboard.type(variant, delay=15)
    page.wait_for_timeout(300)

    log("info", "Execut cautarea variantei")
    try:
        dialog.locator('[title*="Execute" i]').first.click(timeout=6000)
        log("info", "Click Execute trimis in dialogul Get Variant")
    except Exception:
        page.keyboard.press("F8")
        log("warn", "Execute din dialog nu a putut fi apasat; F8 trimis")
    _wait_until_idle(page, timeout_s=6)

    # SAP either applies the variant directly or lists the matches. The result
    # table is not consistent across SAP GUI versions, so do not rely only on
    # ARIA grid roles when locating the requested row.
    picked = dialog_frame.evaluate(
        """(wanted) => {
            const dialogs = Array.from(document.querySelectorAll(
                '[role=dialog], [id^="webguiPopupWindow"], .lsPWNew, .LS_PopupWindow2'
            ))
                .filter(d => d.offsetParent);
            const target = wanted.trim().toLowerCase();
            for (const d of dialogs) {
                const rows = Array.from(d.querySelectorAll('[role=row]'));
                for (const r of rows) {
                    const cells = Array.from(
                        r.querySelectorAll('[role=gridcell],[role=rowheader]')
                    );
                    for (const c of cells) {
                        const text = (c.textContent || '').trim().toLowerCase();
                        if (text === target) {
                            return c.id || r.id || null;
                        }
                    }
                }
            }
            return null;
        }""",
        variant,
    )
    if picked:
        log("info", f"Selectez varianta din lista ({picked})")
        dialog_frame.locator(f"#{_css_escape(picked)}").dblclick(timeout=6000)
    else:
        try:
            dialog.get_by_text(variant, exact=True).last.dblclick(timeout=5000)
            log("info", "Selectez varianta din lista dupa textul vizibil")
        except Exception:
            remaining = _visible_sap_dialog_texts(page)
            _raise_if_missing_variant(page, variant, remaining, log)
            if remaining:
                raise RuntimeError(
                    "Dialogul Get Variant a ramas deschis; varianta nu a fost "
                    f"selectata. Dialoge active: {remaining[:2]}"
                )
            log("info", "Varianta a fost aplicata direct (fara lista de selectie)")

    # A successful double-click can trigger a navigation or an asynchronous
    # refresh. Never touch the selection controls while SAP still owns a modal
    # dialog or reports a busy overlay.
    deadline = time.time() + 20
    while time.time() < deadline:
        dialogs = 0
        busy = False
        for frame in page.frames:
            try:
                state = frame.evaluate(
                    """(selector) => {
                        const visible = el => el && el.offsetParent !== null;
                        return {
                            dialogs: Array.from(document.querySelectorAll(selector))
                                .filter(visible).length,
                            busy: Array.from(document.querySelectorAll(
                                '[aria-busy="true"], .lsBusy, .lsProgress, .sapUiLocalBusyIndicator'
                            )).some(visible),
                        };
                    }""",
                    _VARIANT_DIALOG_SELECTOR,
                )
                dialogs += int(state.get("dialogs") or 0)
                busy = busy or bool(state.get("busy"))
            except Exception:
                continue
        if not dialogs and not busy:
            page.wait_for_timeout(150)
            active = _classify_frame(page, timeout_s=10)
            log(
                "info",
                "Varianta aplicata; ecranul de selectie reconfirmat "
                f"({_frame_debug_label(page, active)})",
            )
            return active
        leftover = _visible_sap_dialog_texts(page)
        _raise_if_missing_variant(page, variant, leftover, log)
        page.wait_for_timeout(250)
    leftover = _visible_sap_dialog_texts(page)
    _raise_if_missing_variant(page, variant, leftover, log)
    raise TimeoutError(
        "SAP nu a inchis dialogul Get Variant sau indicatorul busy dupa "
        "aplicarea variantei."
    )


# Injected into the SAP frame; kept at module level so it can be validated
# against a real browser engine without driving SAP.
_JS_FIND_SCHEME_FIELD = r"""() => {
    const visible = el => el && el.offsetParent !== null;
    const inputs = Array.from(document.querySelectorAll('input'))
        .filter(i => visible(i) && !i.readOnly && !i.disabled &&
                     i.type !== 'checkbox' && i.type !== 'radio');
    const labelFor = (inp) => {
        const parts = [
            inp.getAttribute('title') || '',
            inp.getAttribute('aria-label') || '',
        ];
        const ref = inp.getAttribute('aria-labelledby');
        if (ref) {
            for (const id of ref.split(/\s+/)) {
                const el = document.getElementById(id);
                if (el) parts.push(el.textContent || '');
            }
        }
        return parts.join(' ');
    };
    const matches = inputs
        .filter(i => /numbering\s*scheme|nummernschema/i.test(labelFor(i)))
        .map(i => ({ el: i, left: i.getBoundingClientRect().left }));
    matches.sort((a, b) => a.left - b.left);
    const hit = matches[0];
    if (!hit) {
        return {
            found: false,
            labels: inputs.slice(0, 40).map(labelFor).filter(Boolean),
        };
    }
    return {
        found: true,
        id: hit.el.id,
        current: (hit.el.value || '').trim(),
        title: hit.el.getAttribute('title'),
    };
}"""

_JS_READ_MESSAGES = r"""() => {
    const visible = el => !!el && el.offsetParent !== null;
    const texts = Array.from(document.querySelectorAll(
        '[role=alert],[role=status],#msgarea,#msgarea-itms,#msgpanel,.lsMessageBar'
    )).filter(visible)
      .map(a => (a.textContent || '').replace(/\s+/g, ' ').trim())
      .filter(Boolean);
    const re = /\berror\b|\bfehler\b|make an entry|fill in all required|required entry|mandatory|enter a value|invalid|does not exist|no authorization/i;
    const blocking = texts.filter(t => re.test(t));
    return { texts, blocking };
}"""


def _set_display_all_products(
    page: Page,
    gf: Frame,
    want: bool,
    log: Callable[[str, str], None],
) -> None:
    """Toggle 'Display All Products' so already classified parts can be redone."""
    try:
        state = gf.evaluate(
            """() => {
                const visible = el => !!el && el.offsetParent !== null;
                const box = Array.from(document.querySelectorAll(
                    '[role=checkbox], input[type=checkbox]'
                )).filter(visible).find(el => /Display All Products|Alle Produkte/i.test(
                    (el.getAttribute('aria-label') || '') + ' ' +
                    (el.getAttribute('title') || '') + ' ' +
                    (el.textContent || '')
                ));
                if (!box) return null;
                const checked = box.getAttribute('aria-checked') === 'true' ||
                    box.checked === true;
                return { id: box.id, checked };
            }"""
        )
    except Exception as exc:
        log("warn", f"Nu am putut citi bifa 'Display All Products': {str(exc)[:120]}")
        return
    if not state:
        log("warn", "Bifa 'Display All Products' nu este pe ecranul de selectie.")
        return
    if bool(state.get("checked")) == bool(want):
        log(
            "info",
            f"'Display All Products' este deja {'bifat' if want else 'debifat'}",
        )
        return
    try:
        gf.locator(f"#{_css_escape(state['id'])}").click(timeout=5000)
        _wait_until_idle(page, timeout_s=5)
        log("info", f"'Display All Products' setat pe {want}")
    except Exception as exc:
        log("warn", f"Nu am putut comuta 'Display All Products': {str(exc)[:120]}")


def _set_numbering_scheme(
    page: Page,
    gf: Frame,
    scheme: str,
    log: Callable[[str, str], None],
) -> None:
    """Write the numbering scheme into the selection screen field.

    The variant does not always carry the scheme, and SAP refuses to execute
    without it, so it is set explicitly and read back.
    """
    target = gf.evaluate(_JS_FIND_SCHEME_FIELD)

    if not target.get("found"):
        raise RuntimeError(
            "Nu am gasit campul 'Numbering Scheme' pe ecranul de selectie. "
            f"Campuri vizibile: {target.get('labels')}"
        )

    field_id = target["id"]
    if (target.get("current") or "").upper() == scheme.upper():
        log("info", f"Varianta a setat deja schema {scheme}; nu o mai scriu manual")
        return

    log(
        "info",
        f"Fallback: scriu schema {scheme} in campul "
        f"{target.get('title') or field_id!r} (valoare curenta: {target.get('current')!r})",
    )
    field = gf.locator(f"#{_css_escape(field_id)}")
    field.click(timeout=6000)
    gf.evaluate(
        """(id) => {
            const el = document.getElementById(id);
            if (!el) return false;
            el.focus();
            el.value = '';
            el.setAttribute('value', '');
            return true;
        }""",
        field_id,
    )
    page.keyboard.type(scheme, delay=20)
    page.keyboard.press("Tab")

    read_back = """(id) => {
            const el = document.getElementById(id);
            return el ? (el.value || '').trim() : null;
        }"""
    deadline = time.time() + 3
    written = None
    while time.time() < deadline:
        written = gf.evaluate(read_back, field_id)
        if (written or "").upper() == scheme.upper():
            break
        page.wait_for_timeout(150)
    if (written or "").upper() != scheme.upper():
        raise RuntimeError(
            f"Schema de numerotare nu a fost acceptata de SAP: "
            f"asteptat {scheme!r}, gasit {written!r}"
        )
    log("info", f"Schema de numerotare confirmata: {written}")


def _sap_messages(gf: Frame) -> dict:
    return gf.evaluate(_JS_READ_MESSAGES)


def _classify_report(gf: Frame) -> dict:
    return gf.evaluate(_JS_CLASSIFY_REPORT_STATE)


def _products_from_report(report: dict) -> set[str]:
    products = {
        str(product).strip()
        for product in (report.get("products") or [])
        if str(product).strip()
    }
    blob = "\n".join(
        str(item)
        for item in list(report.get("success") or []) + list(report.get("lines") or [])
        if item
    )
    products.update(_PRODUCT_CLASSIFIED_RE.findall(blob))
    return products


def _log_report_lines(
    report: dict,
    seen: set[str],
    log: Callable[[str, str], None],
) -> None:
    noise = re.compile(
        r"^(Typ|Message text|List|Display logs|\d+)$",
        re.IGNORECASE,
    )
    for line in report.get("lines") or []:
        text = str(line or "").strip()
        if not text or text in seen or noise.match(text):
            continue
        seen.add(text)
        matches = _PRODUCT_CLASSIFIED_RE.findall(text)
        if len(matches) > 1 or len(text) > 180:
            log(
                "info",
                "Raport SAP: bloc concatenat cu "
                f"{len(matches) or 1} produse confirmate",
            )
            continue
        log("info", f"Raport SAP: {text}")


def _acknowledge_classify_report(
    page: Page,
    gf: Frame,
    log: Callable[[str, str], None],
) -> bool:
    """Press the report's explicit Continue button before sending next F8."""
    try:
        marked = gf.evaluate(_JS_MARK_MASS_DIALOG_BUTTON, "confirm")
        if marked.get("found"):
            gf.locator('[data-rpa-mass-btn="1"]').first.click(timeout=5000)
            log(
                "info",
                "Raport SAP confirmat cu butonul Continue "
                f"({marked.get('title') or marked.get('id')})",
            )
            page.wait_for_timeout(700)
            if not _classify_report(gf).get("open"):
                return True
    except Exception as exc:
        log("warn", f"Click real pe Continue SAP a esuat: {str(exc)[:100]}")

    continue_selectors = (
        '[title="Continue (Enter)"]:visible',
        '[aria-label="Continue (Enter)"]:visible',
        "#M1\\:50\\:\\:btn\\[0\\]:visible",
        '[title*="Continue" i]:visible',
        '[aria-label*="Continue" i]:visible',
    )
    for selector in continue_selectors:
        try:
            button = gf.locator(selector).first
            button.click(timeout=5000)
            log("info", "Raport SAP confirmat cu butonul Continue (Enter)")
            page.wait_for_timeout(700)
            if not _classify_report(gf).get("open"):
                return True
        except Exception as exc:
            log(
                "warn",
                f"Continue SAP nu a putut fi apasat ({selector}): {str(exc)[:100]}",
            )

    # Keyboard Enter is only a fallback. F8/Escape must not be used here: in
    # this report SAP assigns F8 to sorting and Escape can leave the report open.
    try:
        page.keyboard.press("Enter")
        log("warn", "Fallback Enter trimis pentru Continue SAP")
    except Exception:
        pass
    page.wait_for_timeout(700)
    return not _classify_report(gf).get("open")


# Export flow JS, kept at module level for the same reason as the scheme
# field lookup: it can be replayed against a plain browser page without SAP.

# Marks the element instead of clicking it, so Playwright can issue a real
# mouse click: SAP ITS menus ignore synthetic el.click().
_JS_MARK_BY_TEXT = r"""({ pattern, roots }) => {
    const re = new RegExp(pattern, 'i');
    const visible = el => el && el.offsetParent !== null;
    document.querySelectorAll('[data-rpa-click]').forEach(
        el => el.removeAttribute('data-rpa-click')
    );
    const all = Array.from(document.querySelectorAll(roots))
        .filter(visible)
        .filter(el => re.test((el.textContent || '').replace(/\s+/g, ' ').trim()));
    // A container holds the whole menu text and matches too; keep only leaves.
    const leaves = all.filter(el => !all.some(o => o !== el && el.contains(o)));
    leaves.sort(
        (a, b) => (a.textContent || '').trim().length - (b.textContent || '').trim().length
    );
    const hit = leaves[0];
    if (!hit) return { found: false, candidates: all.length };
    hit.setAttribute('data-rpa-click', '1');
    return {
        found: true,
        id: hit.id || null,
        tag: hit.tagName,
        cls: String(hit.className || '').slice(0, 80),
        text: (hit.textContent || '').replace(/\s+/g, ' ').trim().slice(0, 60),
    };
}"""

_JS_MARK_EXPORT_TO_BUTTON = r"""() => {
    const visible = el => el && el.offsetParent !== null;
    document.querySelectorAll('[data-rpa-click]').forEach(
        el => el.removeAttribute('data-rpa-click')
    );
    const buttons = Array.from(
        document.querySelectorAll('[role=button], button')
    ).filter(visible);
    const hit = buttons.find(b => {
        const title = (b.getAttribute('title') || '').toLowerCase();
        const text = (b.textContent || '').replace(/\s+/g, ' ').trim().toLowerCase();
        return title.includes('export data') || text.startsWith('export to');
    });
    if (!hit) {
        return {
            found: false,
            buttons: buttons.slice(0, 25).map(
                b => ((b.getAttribute('title') || '') + '|' +
                      (b.textContent || '').replace(/\s+/g, ' ').trim()).slice(0, 50)
            ),
        };
    }
    hit.setAttribute('data-rpa-click', '1');
    return { found: true, id: hit.id || null };
}"""

# Fallback for controls that only react to a full mouse event sequence.
_JS_SYNTH_CLICK_MARKED = r"""() => {
    const el = document.querySelector('[data-rpa-click="1"]');
    if (!el) return false;
    el.scrollIntoView({ block: 'center' });
    for (const type of ['mouseover', 'mousedown', 'mouseup', 'click']) {
        el.dispatchEvent(
            new MouseEvent(type, { bubbles: true, cancelable: true, view: window })
        );
    }
    return true;
}"""

_JS_EXPORT_DIALOG_OPEN = r"""() => {
    const visible = el => el && el.offsetParent !== null;
    const popups = Array.from(
        document.querySelectorAll('[role=dialog], .lsPWNew, .LS_PopupWindow2')
    ).filter(visible);
    const text = popups.map(p => (p.textContent || '').replace(/\s+/g, ' ')).join(' ');
    return {
        open: popups.length > 0,
        isExportAs: /export as|file name for export|destination/i.test(text),
        sample: text.slice(0, 200),
    };
}"""

_JS_MARK_DOWNLOAD_CONFIRM = r"""() => {
    const visible = el => el && el.offsetParent !== null;
    document.querySelectorAll('[data-rpa-download-confirm]').forEach(
        el => el.removeAttribute('data-rpa-download-confirm')
    );

    const mark = (el, via, popup) => {
        if (!el) return null;
        el.setAttribute('data-rpa-download-confirm', '1');
        return {
            found: true,
            via,
            id: el.id || null,
            tag: el.tagName,
            cls: String(el.className || '').slice(0, 100),
            text: (el.textContent || '').replace(/\s+/g, ' ').trim().slice(0, 20),
            popupId: popup && popup.id || null,
            popupCls: popup && String(popup.className || '').slice(0, 100) || null,
        };
    };

    const known = document.getElementById('UpDownDialogChoose');
    if (visible(known)) {
        return mark(known, 'known-id', known.closest('[id^="webguiPopupWindow"]'));
    }

    const popupSelectors = [
        '[id^="webguiPopupWindow"]',
        '[role=dialog]',
        '.lsPWNew',
        '.LS_PopupWindow2',
    ].join(',');
    const popups = Array.from(document.querySelectorAll(popupSelectors))
        .filter(visible)
        .filter(p => /file name|export/i.test(p.textContent || ''));
    for (const popup of popups) {
        const controls = Array.from(
            popup.querySelectorAll('[role=button], button, .lsButton')
        ).filter(visible);
        const ok = controls.find(control => {
            const text = (control.textContent || '').replace(/\s+/g, ' ').trim();
            const accessKey = (control.getAttribute('accesskey') || '').toLowerCase();
            return text.toLowerCase() === 'ok' || accessKey === 'o';
        });
        const result = mark(ok, 'popup-search', popup);
        if (result) return result;
    }
    return { found: false };
}"""

_JS_DOWNLOAD_CONFIRM_STATE = r"""() => {
    const visible = el => el && el.offsetParent !== null;
    const known = document.getElementById('UpDownDialogChoose');
    if (visible(known)) return { open: true, id: known.id };
    const popups = Array.from(document.querySelectorAll(
        '[id^="webguiPopupWindow"], [role=dialog], .lsPWNew, .LS_PopupWindow2'
    )).filter(el => visible(el) && /file name|export/i.test(el.textContent || ''));
    return {
        open: popups.length > 0,
        id: popups[0] && popups[0].id || null,
    };
}"""

_JS_MASS_DIALOG_PROBE = r"""(scheme) => {
    const visible = el => el && el.offsetParent !== null;
    const dialogs = Array.from(document.querySelectorAll(
        '[role="dialog"], [id^="webguiPopupWindow"], .lsPWNew, .LS_PopupWindow2'
    )).filter(visible);
    if (!dialogs.length) return { open: false, relevant: false };

    const readLabel = (inp) => {
        const parts = [
            inp.getAttribute('title') || '',
            inp.getAttribute('aria-label') || '',
        ];
        const refs = (inp.getAttribute('aria-labelledby') || '').split(/\s+/).filter(Boolean);
        for (const id of refs) {
            const node = document.getElementById(id);
            if (node) parts.push(node.textContent || '');
        }
        if (inp.id) {
            const esc = (window.CSS && CSS.escape) ? CSS.escape(inp.id) :
                inp.id.replace(/([:\[\]\.])/g, '\\$1');
            const label = document.querySelector('label[for="' + esc + '"]');
            if (label) parts.push(label.textContent || '');
        }
        return parts.join(' ').replace(/\s+/g, ' ').trim();
    };

    const schemeCode = String(scheme || '').trim().toUpperCase();
    const isSchemeField = (inp) => {
        const label = readLabel(inp).toLowerCase();
        return /numbering\s*scheme|nummernschema|schema/.test(label);
    };
    const isLikelyTariffField = (inp) => {
        const label = readLabel(inp).toLowerCase();
        if (/numbering\s*scheme|nummernschema|schema|valid|from|to|range|product/.test(label)) {
            return false;
        }
        return /tariff|commodity|customs|hs\s*code|classification\s*code|\bcode\b/.test(label);
    };

    const inspect = (dialog) => {
        const text = (dialog.textContent || '').replace(/\s+/g, ' ').trim();
        const inputs = Array.from(dialog.querySelectorAll('input'))
            .filter(i => visible(i) && i.type !== 'checkbox' && i.type !== 'hidden');
        const editable = inputs.filter(i => !i.readOnly && !i.disabled);
        const schemeInput = schemeCode ? inputs.find(
            i => (i.value || '').trim().toUpperCase() === schemeCode
        ) : null;
        const schemeField = inputs.find(isSchemeField) || null;
        const schemeTextMatch = !!(
            schemeCode && text.toUpperCase().includes(schemeCode)
        );
        const labelledTarget = editable.find(i => {
            if (schemeField && i === schemeField) return false;
            return isLikelyTariffField(i);
        }) || null;
        const relevant =
            /mass\s+classif|numbering\s*scheme|tariff|commodity|customs|hs\s*code/i.test(text) ||
            !!schemeField || !!labelledTarget;
        return {
            dialog,
            text,
            inputs,
            editable,
            schemeInput,
            schemeField,
            schemeTextMatch,
            labelledTarget,
            relevant,
        };
    };

    const inspected = dialogs.map(inspect);
    const picked = inspected.slice().reverse().find(x => x.relevant) || null;
    if (!picked) {
        const last = inspected[inspected.length - 1];
        return {
            open: true,
            relevant: false,
            dialogId: last.dialog.id || null,
            sample: last.text.slice(0, 180),
            inputHints: last.inputs.slice(0, 8).map(i => ({
                id: i.id || null,
                label: readLabel(i).slice(0, 60),
                value: String(i.value || '').slice(0, 20),
                ro: !!i.readOnly,
                dis: !!i.disabled,
            })),
        };
    }

    const {
        dialog, text, inputs, editable, schemeInput, schemeField,
        schemeTextMatch, labelledTarget,
    } = picked;
    let target = labelledTarget;
    if (!target) {
        const generic = editable.filter(i => {
            if (schemeField && i === schemeField) return false;
            const label = readLabel(i).toLowerCase();
            return !/numbering\s*scheme|nummernschema|schema|valid|from|to|range|product/.test(label);
        });
        // Generic fallback is allowed only inside a dialog already proven to be
        // Mass Classification; an arbitrary SAP popup must never receive HS data.
        target = generic[0] || null;
    }

    return {
        open: true,
        relevant: true,
        dialogId: dialog.id || null,
        sample: text.slice(0, 180),
        schemeFound: !!schemeInput,
        schemeValue: schemeInput ? (schemeInput.value || '').trim() :
            (schemeField ? (schemeField.value || '').trim() : null),
        schemeFieldId: schemeField ? (schemeField.id || null) : null,
        schemeTextMatch,
        targetId: target ? (target.id || null) : null,
        targetTitle: target ? readLabel(target) : null,
        inputHints: inputs.slice(0, 8).map(i => ({
            id: i.id || null,
            label: readLabel(i).slice(0, 60),
            value: String(i.value || '').slice(0, 20),
            ro: !!i.readOnly,
            dis: !!i.disabled,
        })),
    };
}"""

_JS_MASS_DIALOG_BUTTONS = r"""() => {
    const visible = el => el && el.offsetParent !== null;
    const labelOf = el => ((el.getAttribute('title') || '') + ' ' +
        (el.getAttribute('aria-label') || '') + ' ' +
        (el.textContent || '')).replace(/\s+/g, ' ').trim();
    const kindOf = (el, label) => {
        const low = label.toLowerCase();
        const id = String(el.id || '');
        if (/start\s*mass\s*classif|mass\s*classif.*f8/.test(low)) return 'start';
        if (/cancel|close|abbruch|beenden/.test(low)) return 'cancel';
        if (/continue|enter|Ã¼bernehmen|uebernehmen|weiter|\bok\b/.test(low)) {
            return 'confirm';
        }
        if (/btn\[0\]/.test(id)) return 'confirm';
        return 'other';
    };
    const dialogs = Array.from(document.querySelectorAll(
        '[role="dialog"], [id^="webguiPopupWindow"], .lsPWNew, .LS_PopupWindow2'
    )).filter(visible);
    const dialog = dialogs[dialogs.length - 1];
    if (!dialog) return { open: false, dialogId: null, buttons: [] };

    const nodes = Array.from(dialog.querySelectorAll(
        '[role=button], button, .lsButton, [id*="btn["]'
    )).filter(visible);
    const buttons = nodes.map(el => {
        const label = labelOf(el);
        return {
            id: el.id || null,
            label: label.slice(0, 80),
            kind: kindOf(el, label),
        };
    });
    return {
        open: true,
        dialogId: dialog.id || null,
        buttons,
    };
}"""

_JS_MARK_MASS_DIALOG_BUTTON = r"""(kind) => {
    const wanted = String(kind || '').toLowerCase();
    const visible = el => !!el && el.offsetParent !== null;
    document.querySelectorAll('[data-rpa-mass-btn]').forEach(
        el => el.removeAttribute('data-rpa-mass-btn')
    );
    const labelOf = el => ((el.getAttribute('title') || '') + ' ' +
        (el.getAttribute('aria-label') || '') + ' ' +
        (el.textContent || '')).replace(/\s+/g, ' ').trim();
    const controlSelector =
        '[role=button], button, .lsButton, [id*="btn["], [title], [aria-label]';
    const dialogs = Array.from(document.querySelectorAll(
        '[role="dialog"], [id^="webguiPopupWindow"], .lsPWNew, .LS_PopupWindow2'
    )).filter(visible).reverse();
    const nodes = wanted === 'start'
        ? Array.from(document.querySelectorAll(controlSelector)).filter(visible)
        : dialogs.flatMap(dialog =>
            Array.from(dialog.querySelectorAll(controlSelector)).filter(visible)
        );
    const scored = nodes.map(el => {
        const label = labelOf(el);
        const low = label.toLowerCase();
        const id = String(el.id || '');
        let score = 0;
        if (wanted === 'start') {
            if (/start\s*mass\s*classif/.test(low)) score = 3;
            else if (/mass\s*classif.*f8|\bf8\b.*classif/.test(low)) score = 2;
            else if (/toolbar_btn8|btn\[8\]/.test(id) && /classif/.test(low)) score = 1;
            else if (/toolbar_btn8|btn\[8\]/.test(id) && /f8|start|execute/.test(low)) score = 1;
        } else if (wanted === 'confirm') {
            if (/start\s*mass\s*classif|cancel|close|abbruch/.test(low)) score = 0;
            else if (/continue \(enter\)|continue/.test(low) || /btn\[0\]/.test(id)) score = 2;
        } else if (wanted === 'cancel') {
            if (/cancel|close|abbruch|beenden/.test(low)) score = 2;
        }
        return { el, label, id, score };
    }).filter(x => x.score > 0);
    scored.sort((a, b) => b.score - a.score);
    const hit = scored[0];
    if (!hit) {
        return {
            found: false,
            reason: 'no-button',
            buttons: nodes.slice(0, 20).map(el => (
                (el.getAttribute('title') || el.id || '')
            ).slice(0, 60)),
        };
    }
    try { hit.el.scrollIntoView({ block: 'center', inline: 'center' }); } catch (e) {}
    hit.el.setAttribute('data-rpa-mass-btn', '1');
    return {
        found: true,
        id: hit.id || null,
        title: hit.label.slice(0, 80),
        score: hit.score,
    };
}"""

_JS_SYNTH_CLICK_MASS_MARKED = r"""() => {
    const el = document.querySelector('[data-rpa-mass-btn="1"]');
    if (!el) return false;
    try { el.scrollIntoView({ block: 'center', inline: 'center' }); } catch (e) {}
    for (const type of ['mouseover', 'mousedown', 'mouseup', 'click']) {
        el.dispatchEvent(
            new MouseEvent(type, { bubbles: true, cancelable: true, view: window })
        );
    }
    return true;
}"""

_JS_MASS_DIALOG_DIAGNOSTICS = r"""() => {
    const visible = el => !!el && el.offsetParent !== null;
    const get = (id) => {
        const el = document.getElementById(id);
        if (!el) return { id, exists: false, visible: false, title: '' };
        return {
            id,
            exists: true,
            visible: visible(el),
            title: (el.getAttribute('title') || el.getAttribute('aria-label') || '').slice(0, 80),
            text: (el.textContent || '').replace(/\s+/g, ' ').trim().slice(0, 80),
        };
    };
    return {
        startCandidates: [get('M1:48::btn[8]'), get('M0:50::btn[8]'), get('M1:50::btn[8]'), get('C640_toolbar_btn8')],
        container37: get('M1:37-itms'),
        container46: get('M1:46'),
    };
}"""

_JS_CLASSIFY_REPORT_STATE = r"""() => {
    const visible = el => el && el.offsetParent !== null;
    const popups = Array.from(document.querySelectorAll(
        '[id^="SAPMSSY"], [id^="webguiPopupWindow"], [role="dialog"], .lsPWNew, .LS_PopupWindow2'
    )).filter(visible);

    const popup = popups.find(p => {
        const id = (p.id || '').toUpperCase();
        if (id.startsWith('SAPMSSY')) return true;
        if (p.querySelector('[id^="userarealist"]')) return true;
        const txt = (p.textContent || '');
        return /message\s*text|classified\s+with\s+number|is\s+classified/i.test(txt);
    }) || null;

    if (!popup) return {
        open: false, popupId: null, lines: [], success: [], errors: [],
        products: [], footerCount: 0, visibleCount: 0, itemCount: 0
    };

    const list = popup.querySelector('[id^="userarealist"]') ||
        popup.querySelector('.lsAbapList') || popup;
    const itemNodes = Array.from(list.querySelectorAll(
        '.lsAbapList__item, [role=row], tr'
    ));
    const rawNodes = Array.from(popup.querySelectorAll(
        '[id^="userarealist"] .lsAbapList__item, .lsAbapList__item, '
        + '[id^="userarealist"] [role=row], [id^="userarealist"] tr, '
        + '[role=text], [role=alert], [role=status]'
    ))
        .map(n => (n.textContent || '').replace(/\s+/g, ' ').trim())
        .filter(Boolean);
    const rawText = String(
        (popup.innerText || '') + '\n' + (popup.textContent || '')
    )
        .split(/\r?\n/)
        .map(t => t.replace(/\s+/g, ' ').trim())
        .filter(Boolean);
    const raw = rawNodes.concat(rawText);

    const seen = new Set();
    const lines = [];
    for (const t of raw) {
        if (!seen.has(t)) {
            seen.add(t);
            lines.push(t);
        }
    }

    const blob = lines.join('\n');
    const productRe = /Product\s+([0-9A-Za-z._/-]+)\s+(?:is\s+)?classified/gi;
    const products = [];
    const seenProducts = new Set();
    let match;
    while ((match = productRe.exec(blob))) {
        if (!seenProducts.has(match[1])) {
            seenProducts.add(match[1]);
            products.push(match[1]);
        }
    }

    const success = lines.filter(t =>
        /is\s+classified\s+with\s+number|classified\s+with\s+number|saved\s+successfully|erfolgreich\s+gesichert|wurde\s+gesichert/i.test(t)
    );
    const errors = lines.filter(t => {
        if (/\b0\s+(?:errors?|fehler|failed|blocked)\b/i.test(t)) return false;
        return /\berror\b|\bfehler\b|not\s+classified|failed|blocked|not\s+saved|invalid|does\s+not\s+exist|no\s+authorization/i.test(t);
    });
    const footerNums = lines
        .filter(t => /^\d+$/.test(t))
        .map(t => Number(t));
    const footerCount = footerNums.length ? footerNums[footerNums.length - 1] : 0;

    return {
        open: true,
        popupId: popup.id || null,
        lines: lines.slice(0, 80),
        success,
        errors,
        products,
        footerCount,
        visibleCount: products.length,
        itemCount: itemNodes.length,
    };
}"""

_JS_SCROLL_CLASSIFY_REPORT = r"""() => {
    const visible = el => el && el.offsetParent !== null;
    const popups = Array.from(document.querySelectorAll(
        '[id^="SAPMSSY"], [id^="webguiPopupWindow"], [role="dialog"], .lsPWNew, .LS_PopupWindow2'
    )).filter(visible);
    const popup = popups.find(p => {
        const id = (p.id || '').toUpperCase();
        if (id.startsWith('SAPMSSY')) return true;
        if (p.querySelector('[id^="userarealist"]')) return true;
        const txt = (p.textContent || '');
        return /message\s*text|classified\s+with\s+number|is\s+classified/i.test(txt);
    }) || null;
    if (!popup) return { scrolled: false, reason: 'no-popup' };

    const list = popup.querySelector('[id^="userarealist"]') ||
        popup.querySelector('.lsAbapList') || popup;
    const items = Array.from(list.querySelectorAll(
        '.lsAbapList__item, [role=row], tr'
    ));
    const last = items[items.length - 1] || null;
    const beforeTop = list.scrollTop;
    const beforeHeight = list.scrollHeight;
    try {
        if (last && last.scrollIntoView) last.scrollIntoView({ block: 'end' });
    } catch (e) {}
    const candidates = [list, popup, list.parentElement].filter(Boolean);
    for (const el of candidates) {
        if (el.scrollHeight > el.clientHeight + 4) {
            el.scrollTop = Math.min(
                el.scrollTop + Math.max(el.clientHeight, 48),
                el.scrollHeight
            );
        }
    }
    return {
        scrolled: list.scrollTop !== beforeTop || list.scrollHeight !== beforeHeight,
        top: list.scrollTop,
        height: list.scrollHeight,
        client: list.clientHeight,
        items: items.length,
    };
}"""

_JS_MARK_PRODUCT_MULTI_BUTTON = r"""() => {
    const visible = el => el && el.offsetParent !== null;
    document.querySelectorAll('[data-rpa-click]').forEach(
        el => el.removeAttribute('data-rpa-click')
    );

    const products = Array.from(document.querySelectorAll(
        'input[title="Product"], input[aria-label="Product"]'
    )).filter(visible);
    products.sort((a, b) =>
        a.getBoundingClientRect().left - b.getBoundingClientRect().left
    );
    const product = products[0] || null;

    const controls = Array.from(document.querySelectorAll(
        '[role="button"], button, div[ct="B"]'
    )).filter(visible);

    const score = (el) => {
        const label = (
            (el.getAttribute('title') || '') + ' ' +
            (el.getAttribute('aria-label') || '') + ' ' +
            (el.textContent || '')
        ).toLowerCase();
        const lsdata = (el.getAttribute('lsdata') || '').toLowerCase();
        let s = 0;
        if (/multiple\s*selection|multiple\s*values|mehrfachselektion|mehrfachwerte/.test(label)) s += 120;
        if (/product|material/.test(label)) s += 40;
        if (/valu_push|multiple|multi/.test(lsdata)) s += 30;
        if (product) {
            const a = product.getBoundingClientRect();
            const b = el.getBoundingClientRect();
            const dx = Math.abs((a.left + a.width) - b.left);
            const dy = Math.abs(a.top - b.top);
            s += Math.max(0, 30 - Math.min(30, Math.floor(dx / 20)));
            s += Math.max(0, 10 - Math.min(10, Math.floor(dy / 18)));
        }
        return s;
    };

    let best = null;
    let bestScore = -1;
    for (const el of controls) {
        const s = score(el);
        if (s > bestScore) {
            best = el;
            bestScore = s;
        }
    }
    if (!best || bestScore < 40) {
        return { found: false, bestScore };
    }
    best.setAttribute('data-rpa-click', '1');
    return {
        found: true,
        id: best.id || null,
        score: bestScore,
        label: (
            (best.getAttribute('title') || '') + ' ' +
            (best.getAttribute('aria-label') || '') + ' ' +
            (best.textContent || '')
        ).replace(/\s+/g, ' ').trim().slice(0, 90),
    };
}"""

_JS_MULTI_SELECTION_POPUP_STATE = r"""() => {
    const visible = el => el && el.offsetParent !== null;
    const popups = Array.from(document.querySelectorAll(
        '[id^="SAPLALDB"], [id^="webguiPopupWindow"], [role="dialog"], .lsPWNew, .LS_PopupWindow2'
    )).filter(visible);

    const relevant = popups.filter(p => {
        const text = (p.textContent || '').replace(/\s+/g, ' ');
        return /Multiple Selection for|Select Single Values|Mehrfachselektion/i.test(text);
    });
    const popup = relevant[relevant.length - 1] || null;
    if (!popup) return { open: false };

    document.querySelectorAll('[data-rpa-multi-input]').forEach(
        el => el.removeAttribute('data-rpa-multi-input')
    );
    const inputs = Array.from(
        popup.querySelectorAll('input[type=text], input:not([type]), textarea')
    ).filter(visible).filter(i => !i.readOnly && !i.disabled);
    const first = inputs[0] || null;
    if (first) first.setAttribute('data-rpa-multi-input', '1');

    return {
        open: true,
        popupId: popup.id || null,
        inputId: first ? first.id || null : null,
        hasInput: !!first,
        filled: inputs.filter(i => String(i.value || '').trim()).length,
        values: inputs.map(i => String(i.value || '').trim()).filter(Boolean).slice(0, 20),
        inputs: inputs.length,
        sample: (popup.textContent || '').replace(/\s+/g, ' ').trim().slice(0, 120),
    };
}"""


def _execute_worklist(page: Page, gf: Frame, log: Callable[[str, str], None]) -> None:
    log("info", "Execut worklistul (F8)")
    try:
        gf.locator('[title*="Execute" i]').first.click(timeout=6000)
    except Exception:
        page.keyboard.press("F8")
    deadline = time.time() + 120
    while time.time() < deadline:
        state = gf.evaluate(
            """() => {
                const visible = el => !!el && el.offsetParent !== null;
                const grid = Array.from(document.querySelectorAll('table, [role=grid]'))
                    .find(g => visible(g) && /Product/i.test(g.textContent || ''));
                const exportBtn = Array.from(document.querySelectorAll(
                    '[title="Export"], [title*="Export" i]'
                )).some(visible);
                const classify = Array.from(document.querySelectorAll(
                    '[title*="Classify Products" i]'
                )).some(visible);
                const txt = (document.body.innerText || '');
                const noData = /No data (was )?found|Keine Daten gefunden/i.test(txt);
                return { hasGrid: !!grid, exportBtn, classify, noData };
            }"""
        )
        if state.get("noData"):
            raise HsEmptyWorklistError(
                "Varianta nu a returnat niciun rand (worklist gol)."
            )
        if state.get("classify") or (state.get("hasGrid") and state.get("exportBtn")):
            log("info", "Worklist afisat")
            return
        # SAP blocks on the selection screen with a status message instead of
        # navigating; surface it now rather than waiting out the timeout.
        messages = _sap_messages(gf)
        if messages.get("blocking"):
            raise RuntimeError(
                "SAP a blocat executia: " + " | ".join(messages["blocking"][:3])
            )
        page.wait_for_timeout(300)
    raise TimeoutError("Worklistul nu a aparut in 120s dupa Execute.")


def _return_to_selection(
    page: Page,
    gf: Frame,
    log: Callable[[str, str], None],
    timeout_s: float = 30.0,
) -> Frame:
    """Navigate back to the selection screen so the next variant can be applied.

    Not verified against a live SAP session; validate on the first real queue
    run and adjust the recovery keys if the screen does not return.
    """
    log("info", "Revin la ecranul de selectie")
    deadline = time.time() + timeout_s
    while time.time() < deadline:
        state = gf.evaluate(
            """() => {
                const visible = el => !!el && el.offsetParent !== null;
                return {
                    hasGetVariant: Array.from(document.querySelectorAll(
                        '[title*="Get Variant" i]'
                    )).some(visible),
                    hasDialog: Array.from(document.querySelectorAll(
                        '[role=dialog], [id^="webguiPopupWindow"], .lsPWNew, .LS_PopupWindow2'
                    )).some(visible),
                };
            }"""
        )
        if state.get("hasGetVariant") and not state.get("hasDialog"):
            return gf
        if state.get("hasDialog"):
            page.keyboard.press("Escape")
        else:
            page.keyboard.press("F3")
        page.wait_for_timeout(1000)
        try:
            gf = _classify_frame(page, timeout_s=3)
        except TimeoutError:
            pass
    raise TimeoutError(
        "Nu am putut reveni la ecranul de selectie Classify Products dupa "
        f"{timeout_s:.0f}s."
    )


def _click_marked_element(
    page: Page,
    gf: Frame,
    log: Callable[[str, str], None],
    label: str,
) -> bool:
    """Click the element previously tagged with data-rpa-click."""
    try:
        gf.locator('[data-rpa-click="1"]').first.click(timeout=6000)
        return True
    except Exception as exc:
        log(
            "warn",
            f"{label}: click real a esuat ({str(exc)[:80]}), incerc evenimente mouse",
        )
    try:
        if gf.evaluate(_JS_SYNTH_CLICK_MARKED):
            return True
    except Exception as exc:
        log("warn", f"{label}: secventa de evenimente a esuat ({str(exc)[:80]})")
    return False


def _select_export_spreadsheet(
    page: Page,
    gf: Frame,
    log: Callable[[str, str], None],
) -> None:
    """Open the toolbar Export menu and pick the Spreadsheet entry."""
    gf.locator('[title="Export"], [title*="Export" i]').first.click(timeout=8000)
    page.wait_for_timeout(900)

    marked = gf.evaluate(
        _JS_MARK_BY_TEXT,
        {
            "pattern": r"^(spreadsheet|tabellenkalkulation)",
            "roots": "[role=menuitem], [role=option], .lsMenuItem, li, td, div, span, a",
        },
    )
    if not marked.get("found"):
        raise RuntimeError(
            "Nu am gasit optiunea 'Spreadsheet' in meniul Export "
            f"(candidati: {marked.get('candidates')})."
        )
    log(
        "info",
        f"Optiune 'Spreadsheet' localizata: <{marked.get('tag')}> "
        f"id={marked.get('id')!r} text={marked.get('text')!r}",
    )
    if not _click_marked_element(page, gf, log, "Spreadsheet"):
        raise RuntimeError("Nu am putut apasa optiunea 'Spreadsheet'.")

    # Confirm the menu actually reacted before hunting for the export button.
    deadline = time.time() + 15
    while time.time() < deadline:
        state = gf.evaluate(_JS_EXPORT_DIALOG_OPEN)
        if state.get("isExportAs"):
            log("info", "Dialogul 'Export As' este deschis")
            return
        page.wait_for_timeout(400)
    log(
        "warn",
        "Dialogul 'Export As' nu a aparut inca; continui si verific butonul de export",
    )


def _mark_download_confirmation(
    page: Page,
    preferred: Frame,
) -> tuple[Frame | None, dict[str, Any]]:
    frames: list[Frame] = []
    seen: set[int] = set()
    for frame in (preferred, *page.frames):
        marker = id(frame)
        if marker in seen:
            continue
        seen.add(marker)
        frames.append(frame)

    for frame in frames:
        try:
            result = frame.evaluate(_JS_MARK_DOWNLOAD_CONFIRM)
        except Exception:
            continue
        if result.get("found"):
            return frame, result
    return None, {}


def _click_download_confirmation(
    page: Page,
    preferred: Frame,
    log: Callable[[str, str], None],
    timeout_s: float = 25.0,
) -> bool:
    """Confirm SAP's final file popup and verify that it actually closed."""
    deadline = time.time() + timeout_s
    keyboard_fallback_used = False

    while time.time() < deadline:
        frame, details = _mark_download_confirmation(page, preferred)
        if frame is None:
            page.wait_for_timeout(300)
            continue

        log(
            "info",
            "Confirmare export localizata: "
            f"frame={frame.name or '<unnamed>'!r}, "
            f"id={details.get('id')!r}, "
            f"class={details.get('cls')!r}, via={details.get('via')!r}, tag={details.get('tag')!r}, text={details.get('text')!r}, popupId={details.get('popupId')!r}",
        )
        control = (
            frame.locator(f"#{details['id']}").first
            if details.get("id")
            else frame.locator('[data-rpa-download-confirm="1"]').first
        )

        if not keyboard_fallback_used:
            try:
                control.scroll_into_view_if_needed(timeout=3000)
                control.click(timeout=6000)
                log("info", "Am trimis click Playwright pe OK")
            except Exception as exc:
                log("warn", f"Clickul Playwright pe OK a esuat: {str(exc)[:120]}")

            page.wait_for_timeout(300)
            keyboard_fallback_used = True

            try:
                state = frame.evaluate(_JS_DOWNLOAD_CONFIRM_STATE)
            except Exception:
                state = {"open": False}
            if not state.get("open"):
                return True

            try:
                control.press("Enter", timeout=3000)
                log("info", "Popup-ul a ramas deschis; am trimis Enter pe OK")
            except Exception as exc:
                log("warn", f"Enter pe controlul OK a esuat: {str(exc)[:120]}")
            try:
                state = frame.evaluate(_JS_DOWNLOAD_CONFIRM_STATE)
            except Exception:
                state = {"open": False}
            if state.get("open"):
                try:
                    control.focus(timeout=2000)
                    page.keyboard.press("Enter")
                    log("info", "Am trimis Enter prin tastatura paginii")
                except Exception as key_exc:
                    log(
                        "warn",
                        f"Enter prin tastatura paginii a esuat: {str(key_exc)[:120]}",
                    )

        try:
            state = frame.evaluate(_JS_DOWNLOAD_CONFIRM_STATE)
        except Exception:
            state = {"open": False}
        if not state.get("open"):
            return True
        page.wait_for_timeout(300)

    return False


def _export_worklist(
    page: Page,
    gf: Frame,
    log: Callable[[str, str], None],
) -> Path:
    """Export -> Spreadsheet -> 'Export As' dialog -> Export to... -> Save dialog.

    SAP ITS renders this as two sequential popups: the first configures file
    name/format/destination (defaults are fine), and only the OK of the second
    (file-save confirmation) popup actually fires the browser download.
    """
    _downloads_dir().mkdir(parents=True, exist_ok=True)
    log("info", "Deschid meniul Export")

    try:
        with page.expect_download(timeout=180000) as dl_info:
            _select_export_spreadsheet(page, gf, log)

            # 'Export As': file name/format/destination already default to
            # xlsx / Local, so only the 'Export to...' button is needed.
            pressed_export_to = False
            deadline = time.time() + 30
            while time.time() < deadline:
                state = gf.evaluate(_JS_MARK_EXPORT_TO_BUTTON)
                if state.get("found"):
                    log("info", f"Apas 'Export to...' (id={state.get('id')!r})")
                    if _click_marked_element(page, gf, log, "Export to..."):
                        pressed_export_to = True
                        break
                page.wait_for_timeout(500)
            if not pressed_export_to:
                probe = gf.evaluate(_JS_MARK_EXPORT_TO_BUTTON)
                log(
                    "warn",
                    "Nu am gasit butonul 'Export to...'; butoane vizibile: "
                    f"{probe.get('buttons')}",
                )

            # Second popup confirms the local file name; its OK fires the download.
            if not _click_download_confirmation(page, gf, log):
                raise TimeoutError(
                    "Popup-ul final de confirmare export a ramas deschis dupa "
                    "click si Enter."
                )
            log("info", "Dialog confirmare descarcare inchis")
        download = dl_info.value
    except PWTimeout as exc:
        try:
            diag = gf.evaluate(_JS_EXPORT_DIALOG_OPEN)
        except Exception:
            diag = {}
        raise TimeoutError(
            "SAP nu a livrat fisierul exportat. Stare popup: "
            f"{diag.get('sample', '')!r}. Foloseste modul semi-automat si "
            "incarca manual exportul."
        ) from exc

    stamp = time.strftime("%Y%m%d_%H%M%S")
    suggested = download.suggested_filename or f"worklist_{stamp}.xlsx"
    suffix = Path(suggested).suffix or ".xlsx"
    target = _downloads_dir() / f"worklist_{stamp}{suffix}"
    download.save_as(str(target))
    log("info", f"Export salvat: {target.name}")
    return target


# ============================================================================
# Product multiple selection + mass classification
# ============================================================================


def _select_all_worklist_rows(
    page: Page,
    gf: Frame,
    log: Callable[[str, str], None],
) -> None:
    """Select all rows in the current worklist grid."""
    try:
        gf.locator(
            '[role=columnheader][title*="select all" i], '
            '[role=columnheader]:has-text("select all"), '
            '[role=columnheader]:has-text("Column for row selection")'
        ).first.click(timeout=5000)
        page.wait_for_timeout(250)
        return
    except Exception as exc:
        log("warn", f"Select all din antet a esuat: {str(exc)[:100]}; incerc Ctrl+A")

    try:
        first_cell = gf.locator(
            "[role=grid] [role=row] [role=gridcell], "
            "table[role=grid] tr td, "
            "table tr td"
        ).first
        first_cell.click(timeout=5000)
        page.keyboard.press("Control+a")
        page.wait_for_timeout(250)
    except Exception as exc:
        raise RuntimeError(
            f"Nu pot selecta toate randurile din worklist: {str(exc)[:120]}"
        ) from exc


def _open_product_multiple_selection(
    page: Page,
    gf: Frame,
    log: Callable[[str, str], None],
) -> Frame:
    """Open Product multiple selection popup from selection screen."""
    marked = gf.evaluate(_JS_MARK_PRODUCT_MULTI_BUTTON)
    if not marked.get("found"):
        raise RuntimeError(
            "Nu am gasit butonul de Multiple Selection pentru Product "
            f"(bestScore={marked.get('bestScore')})."
        )
    log(
        "info",
        f"Buton Multiple Selection localizat: id={marked.get('id')!r}, "
        f"label={marked.get('label')!r}, score={marked.get('score')}",
    )
    if not _click_marked_element(page, gf, log, "Product Multiple Selection"):
        raise RuntimeError("Nu am putut deschide fereastra Product Multiple Selection.")

    deadline = time.time() + 25
    while time.time() < deadline:
        for frame in page.frames:
            try:
                state = frame.evaluate(_JS_MULTI_SELECTION_POPUP_STATE)
            except Exception:
                continue
            if state.get("open"):
                log(
                    "info",
                    "Popup Multiple Selection deschis: "
                    f"id={state.get('popupId')!r}, camp={state.get('inputId')!r}",
                )
                return frame
        page.wait_for_timeout(200)
    raise TimeoutError("Popup-ul Product Multiple Selection nu s-a deschis.")


def _paste_products_via_multiple_selection(
    page: Page,
    gf: Frame,
    products: list[str],
    log: Callable[[str, str], None],
    frame_finder: Callable[..., Frame] | None = None,
) -> Frame:
    """Paste full product list into Product multiple-selection popup and confirm with F8."""
    finder = frame_finder or _classify_frame
    unique = [str(p).strip() for p in dict.fromkeys(products) if str(p).strip()]
    if not unique:
        raise RuntimeError("Lista de produse pentru clasificare este goala.")

    multi_frame = _open_product_multiple_selection(page, gf, log)
    state = multi_frame.evaluate(_JS_MULTI_SELECTION_POPUP_STATE)
    if not state.get("open"):
        raise RuntimeError("Popup Multiple Selection indisponibil dupa deschidere.")

    # Leftovers from a previous group would classify the wrong products.
    if state.get("filled"):
        log(
            "warn",
            f"Popup-ul contine {state.get('filled')} valori vechi; le sterg "
            "(Shift+F4).",
        )
        try:
            multi_frame.locator('[data-rpa-multi-input="1"]').first.click(timeout=5000)
        except Exception:
            pass
        page.keyboard.press("Shift+F4")
        page.wait_for_timeout(700)
        state = multi_frame.evaluate(_JS_MULTI_SELECTION_POPUP_STATE)
        if state.get("filled"):
            raise HsSafetyError(
                "Popup-ul Multiple Selection contine inca valori din grupul "
                f"anterior ({state.get('filled')}); opresc pentru a nu clasifica "
                "produse gresite."
            )
        log("info", "Valorile vechi au fost sterse din popup")

    payload = "\r\n".join(unique)
    _set_windows_clipboard_text(payload)
    log("info", f"Clipboard Windows pregatit cu {len(unique)} produse")

    # SAP reads the clipboard from the page, which the browser refuses while
    # the tab is not the focused document.
    try:
        page.bring_to_front()
    except Exception:
        pass

    # Shift+F12 only reaches SAP when the popup's value list holds the focus.
    if state.get("hasInput"):
        try:
            multi_frame.locator('[data-rpa-multi-input="1"]').first.click(timeout=5000)
        except Exception as exc:
            log("warn", f"Nu pot focaliza lista de valori: {str(exc)[:80]}")
    else:
        log("warn", "Nu am gasit campul 'Single Value' in popup")
    page.wait_for_timeout(200)

    page.keyboard.press("Shift+F12")
    log("info", "Shift+F12 trimis (Upload from Clipboard)")

    after_upload = {"open": False, "filled": 0}
    upload_deadline = time.time() + 8
    while time.time() < upload_deadline:
        try:
            after_upload = multi_frame.evaluate(_JS_MULTI_SELECTION_POPUP_STATE)
        except Exception:
            after_upload = {"open": False, "filled": 0}
        if not after_upload.get("open") or after_upload.get("filled"):
            break
        page.wait_for_timeout(150)

    if not after_upload.get("open"):
        raise RuntimeError(
            "Popup-ul Multiple Selection s-a inchis neasteptat dupa Shift+F12, "
            "inainte de confirmarea Copy (F8)."
        )
    filled = after_upload.get("filled", 0)
    if not filled:
        raise RuntimeError(
            "Shift+F12 nu a incarcat produsele in popup (0 valori). "
            "Verifica permisiunea de clipboard a browserului."
        )
    visible_values = [
        str(v).strip() for v in after_upload.get("values") or [] if str(v).strip()
    ]
    if visible_values and visible_values[0] != unique[0]:
        raise HsSafetyError(
            "Clipboard-ul SAP a incarcat alte valori decat grupul curent: "
            f"asteptat primul produs {unique[0]!r}, gasit {visible_values[0]!r}."
        )
    log("info", f"Valori incarcate in popup: {filled}")

    page.keyboard.press("F8")
    log("info", "F8 trimis (Copy)")

    # Wait until popup is closed.
    deadline = time.time() + 15
    while time.time() < deadline:
        active = finder(page, timeout_s=3)
        try:
            s = active.evaluate(_JS_MULTI_SELECTION_POPUP_STATE)
        except Exception:
            s = {"open": False}
        if not s.get("open"):
            return active
        page.wait_for_timeout(150)
    raise TimeoutError("Popup Multiple Selection a ramas deschis dupa confirmare.")


def _open_mass_classification(
    page: Page,
    gf: Frame,
    log: Callable[[str, str], None],
) -> None:
    """Open Mass Classification dialog using resilient button targeting."""
    selectors = [
        "#C640_toolbar_btn1",
        '[title="Classify Products"]',
        '[title*="Classify Products" i]',
    ]
    attempts: list[str] = []

    for selector in selectors:
        try:
            locator = gf.locator(selector)
            count = min(locator.count(), 4)
        except Exception as exc:
            attempts.append(f"{selector}: {str(exc)[:80]}")
            continue

        for idx in range(count):
            try:
                locator.nth(idx).click(timeout=9000)
                page.wait_for_timeout(900)
                state = gf.evaluate(_JS_MASS_DIALOG_PROBE, "")
                sample = str(state.get("sample") or "")
                if _NO_SELECTION_RE.search(sample):
                    page.keyboard.press("Enter")
                    page.wait_for_timeout(300)
                    raise RuntimeError(
                        "SAP raporteaza 'No lines were selected': bifarea "
                        "randurilor din worklist nu a ajuns la server."
                    )
                if state.get("open") and state.get("relevant"):
                    log(
                        "info",
                        "Dialog Mass Classification deschis: "
                        f"id={state.get('dialogId')!r}",
                    )
                    return
                log(
                    "warn",
                    "Dialogul deschis nu pare cel de clasificare "
                    f"(text={sample[:80]!r}); incerc alt buton.",
                )
                page.keyboard.press("Escape")
                page.wait_for_timeout(350)
            except RuntimeError:
                raise
            except Exception as exc:
                attempts.append(f"{selector}[{idx}]: {str(exc)[:80]}")

    raise RuntimeError(
        "Nu am putut deschide dialogul corect de Mass Classification. "
        f"Incercari: {attempts[:5]}"
    )


def _fill_tariff_code(
    page: Page,
    gf: Frame,
    scheme: str,
    hs_code: str,
    log: Callable[[str, str], None],
) -> None:
    """Verify scheme in dialog (or set it), then type the tariff code."""

    def probe() -> dict[str, Any]:
        return gf.evaluate(_JS_MASS_DIALOG_PROBE, scheme)

    state = probe()
    if not state.get("open"):
        raise RuntimeError("Dialogul 'Mass Classification' nu este deschis.")

    if not state.get("schemeFound") and not state.get("schemeTextMatch"):
        scheme_field_id = state.get("schemeFieldId")
        if scheme_field_id:
            log(
                "warn",
                f"Schema {scheme} nu era precompletata in dialog; o setez explicit.",
            )
            field = gf.locator(f"#{_css_escape(scheme_field_id)}")
            field.click(timeout=6000)
            gf.evaluate(
                """(id) => {
                    const el = document.getElementById(id);
                    if (!el) return false;
                    el.focus();
                    el.value = '';
                    el.setAttribute('value', '');
                    return true;
                }""",
                scheme_field_id,
            )
            page.keyboard.type(scheme, delay=20)
            page.keyboard.press("Tab")
            page.wait_for_timeout(500)
            state = probe()

    if not state.get("schemeFound") and not state.get("schemeTextMatch"):
        raise RuntimeError(
            f"Schema {scheme} nu apare in dialog (id={state.get('dialogId')!r}, "
            f"gasit: {state.get('schemeValue')!r}, inputuri: {state.get('inputHints')}). "
            "Opresc pentru siguranta."
        )

    target_id = state.get("targetId")
    if not target_id:
        raise RuntimeError(
            "Nu am gasit campul pentru codul tarifar in dialog. "
            f"Inputuri: {state.get('inputHints')}"
        )

    log(
        "info",
        f"Camp cod tarifar: {state.get('targetTitle') or target_id} "
        f"(schema confirmata {scheme})",
    )
    field = gf.locator(f"#{_css_escape(target_id)}")
    field.click(timeout=6000)
    gf.evaluate(
        """(id) => {
            const el = document.getElementById(id);
            if (!el) return false;
            el.focus();
            el.value = '';
            el.setAttribute('value', '');
            return true;
        }""",
        target_id,
    )
    page.keyboard.type(hs_code, delay=15)
    page.wait_for_timeout(400)

    typed = gf.evaluate(
        """(id) => {
            const el = document.getElementById(id);
            return el ? (el.value || '').trim() : '';
        }""",
        target_id,
    )
    if typed.replace(" ", "") != hs_code.replace(" ", ""):
        raise RuntimeError(
            f"Codul tarifar nu a fost introdus corect: asteptat {hs_code!r}, gasit {typed!r}"
        )
    log("info", f"Cod tarifar {hs_code} introdus si verificat")

    # SAP ITS keeps the typed value client-side until the field is confirmed.
    # Exactly one Enter belongs to field validation. Starting the mass action is
    # a separate, guarded step and must never happen from this helper.
    gf.evaluate(
        """(id) => {
            const el = document.getElementById(id);
            if (!el) return false;
            el.focus();
            el.dispatchEvent(new Event('input', { bubbles: true }));
            el.dispatchEvent(new Event('change', { bubbles: true }));
            return true;
        }""",
        target_id,
    )
    _confirm_tariff_field(
        page,
        gf,
        target_id,
        scheme,
        hs_code,
        log,
    )


def _confirm_tariff_field(
    page: Page,
    gf: Frame,
    target_id: str,
    scheme: str,
    hs_code: str,
    log: Callable[[str, str], None],
) -> None:
    """Confirm the HS field once and prove the mass dialog is still open."""
    try:
        gf.locator(f"#{_css_escape(target_id)}").focus(timeout=3000)
    except Exception:
        pass
    page.keyboard.press("Enter")
    log("info", "Enter trimis o singura data pentru confirmarea campului HS")
    page.wait_for_timeout(800)

    messages = _sap_messages(gf)
    if messages.get("blocking"):
        raise RuntimeError(
            "SAP a respins codul tarifar: " + " | ".join(messages["blocking"][:3])
        )

    state = gf.evaluate(_JS_MASS_DIALOG_PROBE, scheme)
    if not state.get("open") or not state.get("relevant"):
        raise RuntimeError(
            "Dialogul Mass Classification s-a inchis sau a fost inlocuit dupa "
            "confirmarea campului HS. Opresc pentru a nu raporta fals un commit."
        )
    confirmed_id = state.get("targetId")
    if not confirmed_id:
        raise RuntimeError(
            "Campul codului tarifar nu mai este disponibil dupa Enter. "
            f"Dialog: {state.get('sample')!r}"
        )
    confirmed = gf.evaluate(
        """(id) => {
            const el = document.getElementById(id);
            return el ? (el.value || '').trim() : null;
        }""",
        confirmed_id,
    )
    if (confirmed or "").replace(" ", "") != hs_code.replace(" ", ""):
        raise RuntimeError(
            "SAP nu a pastrat codul tarifar dupa Enter: "
            f"asteptat {hs_code!r}, gasit {confirmed!r}."
        )
    log("info", f"Cod tarifar confirmat de dialog: {confirmed}")


def _start_action_effect(page: Page, gf: Frame, timeout_s: float = 8.0) -> bool:
    """Detect whether Start Mass Classification actually fired."""
    deadline = time.time() + timeout_s
    while time.time() < deadline:
        try:
            report = _classify_report(gf)
            if report.get("open"):
                return True
        except Exception:
            pass
        try:
            probe = gf.evaluate(_JS_MASS_DIALOG_PROBE, "")
            if not probe.get("open") or not probe.get("relevant"):
                return True
        except Exception:
            page.wait_for_timeout(150)
            continue
        try:
            busy = gf.evaluate(
                """() => {
                    const visible = el => !!el && el.offsetParent !== null;
                    return Array.from(document.querySelectorAll(
                        '[aria-busy="true"], .lsBusy, .lsProgress, '
                        + '.sapUiLocalBusyIndicator, [id*="busy" i]'
                    )).some(visible);
                }"""
            )
            if busy:
                return True
        except Exception:
            pass
        page.wait_for_timeout(150)
    return False


def _click_start_mass_classification(
    page: Page,
    gf: Frame,
    log: Callable[[str, str], None],
) -> bool:
    """Deliver one real Start click and return only after a visible SAP effect."""
    try:
        diag = gf.evaluate(_JS_MASS_DIALOG_DIAGNOSTICS)
        log(
            "info",
            "Diag start: "
            + " | ".join(
                f"{c.get('id')} exists={c.get('exists')} visible={c.get('visible')} title={c.get('title')!r}"
                for c in (diag.get("startCandidates") or [])
            ),
        )
    except Exception:
        pass

    marked = {}
    try:
        marked = gf.evaluate(_JS_MARK_MASS_DIALOG_BUTTON, "start")
    except Exception as exc:
        marked = {"found": False, "reason": str(exc)[:80]}

    if marked.get("found"):
        try:
            button = gf.locator('[data-rpa-mass-btn="1"]').first
            button.scroll_into_view_if_needed(timeout=3000)
            button.click(timeout=6000)
            log(
                "info",
                "Click Playwright trimis pe Start Mass Classification "
                f"({marked.get('title') or marked.get('id') or 'marked'})",
            )
            if _start_action_effect(page, gf):
                log("info", "SAP a reactionat la click-ul Start Mass Classification")
                return True
            log(
                "warn",
                "Click-ul Start a fost livrat, dar dialogul nu s-a schimbat; "
                "voi folosi fallback-ul F8 o singura data.",
            )
            return False
        except Exception as exc:
            log("warn", f"Click Start marcat a esuat: {str(exc)[:80]}")
            try:
                if gf.evaluate(_JS_SYNTH_CLICK_MASS_MARKED):
                    log("warn", "Fallback click sintetic trimis pe Start")
                    if _start_action_effect(page, gf):
                        log("info", "SAP a reactionat la fallback-ul sintetic")
                        return True
                    return False
            except Exception as synth_exc:
                log("warn", f"Click sintetic Start esuat: {str(synth_exc)[:80]}")

    if not marked.get("found") and marked.get("buttons"):
        sample = " | ".join(str(x) for x in (marked.get("buttons") or [])[:8])
        log("warn", f"Start Mass Classification negasit. Candidati: {sample}")

    for selector in (
        '[title="Start Mass Classification (F8)"]',
        '[aria-label="Start Mass Classification (F8)"]',
        '[title*="Start Mass Classification" i]',
        '[aria-label*="Start Mass Classification" i]',
        "#C640_toolbar_btn8",
        "#M1\\:48\\:\\:btn\\[8\\]",
        "#M0\\:50\\:\\:btn\\[8\\]",
        "#M1\\:50\\:\\:btn\\[8\\]",
    ):
        try:
            button = gf.locator(selector).first
            if button.count() == 0:
                continue
            button.scroll_into_view_if_needed(timeout=2500)
            button.click(timeout=5000)
            log("info", f"Click Playwright trimis pe Start ({selector})")
            if _start_action_effect(page, gf):
                log("info", "SAP a reactionat la Start Mass Classification")
                return True
            log("warn", "Start a ramas fara efect vizibil dupa click")
            return False
        except Exception as exc:
            log("warn", f"Click Start esuat ({selector}): {str(exc)[:80]}")
    return False


def _close_dialog(page: Page, gf: Frame) -> None:
    marked = {}
    try:
        marked = gf.evaluate(_JS_MARK_MASS_DIALOG_BUTTON, "cancel")
    except Exception:
        pass

    if marked.get("found"):
        try:
            gf.locator('[data-rpa-mass-btn="1"]').first.click(timeout=5000)
            deadline = time.time() + 5
            while time.time() < deadline:
                state = gf.evaluate(_JS_MASS_DIALOG_PROBE, "")
                if not state.get("open") or not state.get("relevant"):
                    return
                page.wait_for_timeout(150)
        except Exception:
            pass

    # Keyboard Escape is only a fallback when a real Cancel click was missing
    # or had no visible effect. Never send both actions unconditionally.
    page.keyboard.press("Escape")
    page.wait_for_timeout(600)


def _commit_classification(
    page: Page,
    gf: Frame,
    log: Callable[[str, str], None],
    expected_products: list[str] | None = None,
) -> set[str]:
    """Start mass classification once, then require explicit SAP confirmation."""
    seen_report_lines: set[str] = set()
    committed_products: set[str] = set()
    expected = {
        str(product).strip()
        for product in (expected_products or [])
        if str(product).strip()
    }

    def collect_report(report: dict) -> None:
        _log_report_lines(report, seen_report_lines, log)
        committed_products.update(_products_from_report(report))
        unexpected = committed_products - expected
        if unexpected:
            raise HsSafetyError(
                "Raportul SAP contine produse care nu apartin grupului curent: "
                + ", ".join(sorted(unexpected)[:10])
            )

    def harvest_open_report() -> dict:
        """Read every classified product, scrolling a virtualized SAP list."""
        last_count = -1
        stagnant = 0
        report: dict = {"open": False}
        for pass_idx in range(1, 41):
            report = _classify_report(gf)
            if not report.get("open"):
                return report
            collect_report(report)
            if report.get("errors"):
                raise RuntimeError(
                    "Raport SAP cu erori: " + " | ".join(report["errors"][:3])
                )
            footer = int(report.get("footerCount") or 0)
            harvested = len(committed_products)
            needed = len(expected)
            log(
                "info",
                "Raport SAP extras: "
                f"confirmate={harvested}/{needed}, footer={footer}, "
                f"vizibile={int(report.get('visibleCount') or 0)}, "
                f"randuri={int(report.get('itemCount') or 0)}, "
                f"trecere={pass_idx}",
            )
            complete = harvested >= needed and (not footer or harvested >= footer)
            if complete:
                return report
            if harvested == last_count:
                stagnant += 1
            else:
                stagnant = 0
                last_count = harvested
            if stagnant >= 3:
                missing = sorted(expected - committed_products)
                log(
                    "warn",
                    "Raportul SAP nu a mai afisat produse noi dupa derulare. "
                    f"Lipsesc {len(missing)}: {', '.join(missing[:8])}",
                )
                return report
            try:
                gf.evaluate(_JS_SCROLL_CLASSIFY_REPORT)
            except Exception as exc:
                log("warn", f"Derularea raportului SAP a esuat: {str(exc)[:80]}")
                return report
            try:
                page.keyboard.press("PageDown")
            except Exception:
                pass
            page.wait_for_timeout(180)
        return report

    initial = gf.evaluate(_JS_MASS_DIALOG_PROBE, "")
    if not initial.get("open") or not initial.get("relevant"):
        raise RuntimeError(
            "Dialogul Mass Classification nu mai este deschis inainte de commit."
        )

    log("info", "Pornesc mentinerea prin controlul Start Mass Classification")
    action_started = _click_start_mass_classification(page, gf, log)
    used_f8_fallback = False

    if not action_started:
        messages = _sap_messages(gf)
        if messages.get("blocking"):
            raise RuntimeError(
                "SAP a raportat eroare la Start: "
                + " | ".join(messages["blocking"][:3])
            )
        report = _classify_report(gf)
        if report.get("open"):
            action_started = True
        else:
            probe = gf.evaluate(_JS_MASS_DIALOG_PROBE, "")
            if probe.get("open") and probe.get("relevant"):
                # A focused text input consumes function keys in SAP ITS. Move
                # focus away, then use the documented F8 action exactly once.
                try:
                    gf.evaluate(
                        """() => {
                            const el = document.activeElement;
                            if (el && typeof el.blur === 'function') el.blur();
                        }"""
                    )
                except Exception:
                    pass
                page.wait_for_timeout(250)
                page.keyboard.press("F8")
                used_f8_fallback = True
                log(
                    "warn",
                    "Click-ul Start nu a produs efect; F8 trimis o singura data "
                    "dupa scoaterea focusului din camp.",
                )
                action_started = _start_action_effect(page, gf, timeout_s=12.0)

    if not action_started:
        probe = gf.evaluate(_JS_MASS_DIALOG_PROBE, "")
        buttons = gf.evaluate(_JS_MASS_DIALOG_BUTTONS)
        raise RuntimeError(
            "SAP nu a reactionat nici la click-ul Start, nici la fallback-ul F8. "
            f"Dialog={probe.get('dialogId')!r}, butoane={buttons.get('buttons')}."
        )

    success_re = re.compile(
        r"saved\s+successfully|data\s+saved|successfully\s+classified|"
        r"is\s+classified|wurde\s+gesichert|erfolgreich\s+gesichert",
        re.IGNORECASE,
    )
    deadline = time.time() + 90
    last_msg = ""
    while time.time() < deadline:
        msgs = _sap_messages(gf)
        texts = [t for t in (msgs.get("texts") or []) if t]
        if texts:
            last_msg = " | ".join(texts[:3])
        if msgs.get("blocking"):
            raise RuntimeError("SAP a raportat eroare: " + " | ".join(msgs["blocking"]))
        if any(success_re.search(t) for t in texts) and not _classify_report(gf).get(
            "open"
        ):
            log("info", "SAP a confirmat clasificarea din mesajele de status")
            if committed_products:
                return committed_products
            if expected:
                log(
                    "warn",
                    "Mesaj SAP de succes fara lista de produse; "
                    "nu marchez grupul ca mentinut complet.",
                )
                return set()
            return set()

        report = _classify_report(gf)
        if report.get("open"):
            report = harvest_open_report()
            missing = sorted(expected - committed_products)
            extra = sorted(committed_products - expected)
            log(
                "info",
                "Confirmare SAP din raport: "
                f"confirmate={len(committed_products)}/{len(expected)}, "
                f"lipsesc={len(missing)}"
                + (f" ({', '.join(missing[:8])})" if missing else ""),
            )
            if extra:
                raise HsSafetyError(
                    "Raportul SAP contine produse care nu apartin grupului curent: "
                    + ", ".join(extra[:10])
                )
            if not _acknowledge_classify_report(page, gf, log):
                raise TimeoutError("Raportul SAP ramane deschis dupa Continue (Enter).")
            if committed_products or report.get("success"):
                log("info", "SAP confirma clasificarea prin raportul final")
                return committed_products
            page.wait_for_timeout(350)
            continue

        page.wait_for_timeout(350)

    if committed_products:
        return committed_products
    raise TimeoutError(
        "Fara confirmare SAP dupa Start Mass Classification"
        + (" si fallback F8" if used_f8_fallback else "")
        + f". Ultimele mesaje SAP: {last_msg!r}. "
        f"Raport: {list(seen_report_lines)[:3]}"
    )


def _classify_group_via_multiple_selection(
    page: Page,
    gf: Frame,
    products: list[str],
    scheme: str,
    variant: str,
    hs_code: str,
    allow_commit: bool,
    cancel: threading.Event,
    log: Callable[[str, str], None],
    display_all_products: bool = False,
) -> tuple[int, Frame, set[str]]:
    """Maintain one HS group: apply the variant, paste its products, classify."""
    if cancel.is_set():
        raise RuntimeError("Anulat de utilizator inainte de clasificare.")

    active_gf = _return_to_selection(page, gf, log, timeout_s=40)
    if variant:
        active_gf = _apply_variant(page, active_gf, variant, log)
    _set_numbering_scheme(page, active_gf, scheme, log)
    _set_display_all_products(page, active_gf, display_all_products, log)
    active_gf = _paste_products_via_multiple_selection(page, active_gf, products, log)

    _execute_worklist(page, active_gf, log)
    active_gf = _classify_frame(page, timeout_s=10)
    _select_all_worklist_rows(page, active_gf, log)

    _open_mass_classification(page, active_gf, log)
    _fill_tariff_code(page, active_gf, scheme, hs_code, log)

    committed_products: set[str] = set()
    if allow_commit:
        committed_products = _commit_classification(page, active_gf, log, products)
    else:
        _close_dialog(page, active_gf)

    log("info", f"[{hs_code}] grup finalizat pentru {len(products)} produse")
    return len(products), active_gf, committed_products


_BLOCKED_PREVIEW_LIMIT = 50


def _account_blocked_materials(
    summary: dict,
    blocked: list[dict],
    scheme: str,
    log: Callable[[str, str], None],
) -> None:
    """Count materials the analysis refused, with one log line instead of many."""
    counted = 0
    for item in blocked or []:
        product = str(item.get("product") or "").strip()
        if not product:
            continue
        reason = str(item.get("reason") or "blocked").strip() or "blocked"
        summary["materials_blocked"] += 1
        counted += 1
        if len(summary["blocked_products"]) < _BLOCKED_PREVIEW_LIMIT:
            summary["blocked_products"].append(product)
        summary["blocked_by_reason"][reason] = (
            summary["blocked_by_reason"].get(reason, 0) + 1
        )
    if not counted:
        return
    reasons = ", ".join(
        f"{REASON_LABELS.get(reason, reason)}={count}"
        for reason, count in sorted(summary["blocked_by_reason"].items())
    )
    log(
        "warn",
        f"[{scheme}] {counted} materiale blocate de precontrol ({reasons}).",
    )


def _classify_groups(
    page: Page,
    gf: Frame,
    groups: list[dict],
    scheme: str,
    variant: str,
    allow_commit: bool,
    cancel: threading.Event,
    cb: Callable[[dict], None],
    log: Callable[[str, str], None],
    blocked: list[dict] | None = None,
    display_all_products: bool = False,
) -> dict:
    """Maintain every approved HS group through the Multiple Selection flow.

    Shared by the manual approval flow (_do_classify) and the automatic
    multi-scheme queue (_do_queue).
    """
    summary = {
        "groups": len(groups),
        "committed": 0,
        "dry_run": 0,
        "failed": 0,
        "skipped": 0,
        "products_selected": 0,
        "materials_committed": 0,
        "materials_dry_run": 0,
        "materials_failed": 0,
        "materials_blocked": 0,
        "blocked_by_reason": {},
        "committed_products": [],
        "dry_run_products": [],
        "failed_products": [],
        "blocked_products": [],
        "safety_aborted": False,
        "abort_reason": "",
    }

    _account_blocked_materials(summary, blocked or [], scheme, log)
    for item in blocked or []:
        product = str(item.get("product") or "").strip()
        if not product:
            continue
        reason = str(item.get("reason") or "blocked").strip() or "blocked"
        reason_label = REASON_LABELS.get(reason, reason)
        cb(
            {
                "type": "material-end",
                "status": "blocked",
                "scheme": str(item.get("scheme") or scheme),
                "product": product,
                "material": product,
                "hs_code": "",
                "text": str(item.get("text") or ""),
                "msg": reason_label,
                "reason": reason,
                "reason_label": reason_label,
            }
        )

    for idx, group in enumerate(groups, start=1):
        if cancel.is_set():
            summary["skipped"] = len(groups) - idx + 1
            log("warn", f"Anulat de utilizator. Ramase: {summary['skipped']}")
            break

        hs_code = group["hs_code"]
        product_rows = [p for p in (group.get("products") or []) if p.get("product")]
        products = [p["product"] for p in product_rows]
        cb(
            {
                "type": "group-start",
                "index": idx,
                "total": len(groups),
                "hs_code": hs_code,
                "count": len(products),
            }
        )
        try:
            selected, gf, committed_products = _classify_group_via_multiple_selection(
                page,
                gf,
                products,
                scheme,
                variant,
                hs_code,
                allow_commit,
                cancel,
                log,
                display_all_products,
            )
            summary["products_selected"] += selected
            if allow_commit:
                missing_count = len(product_rows) - len(committed_products)
                if committed_products and not missing_count:
                    summary["committed"] += 1
                    status = "committed"
                    group_ok = True
                else:
                    summary["failed"] += 1
                    status = "partial" if committed_products else "failed"
                    group_ok = False
                summary["materials_committed"] += len(committed_products)
                summary["materials_failed"] += missing_count
            else:
                summary["dry_run"] += 1
                status = "dry_run"
                group_ok = True
                summary["materials_dry_run"] += len(product_rows)

            for material in product_rows:
                product = str(material.get("product") or "").strip()
                if not product:
                    continue
                material_status = (
                    "committed"
                    if not allow_commit or product in committed_products
                    else "failed"
                )
                if allow_commit and material_status == "committed":
                    summary["committed_products"].append(product)
                elif allow_commit:
                    summary["failed_products"].append(product)
                else:
                    summary["dry_run_products"].append(product)
                cb(
                    {
                        "type": "material-end",
                        "status": material_status if allow_commit else status,
                        "scheme": scheme,
                        "product": product,
                        "material": product,
                        "hs_code": hs_code,
                        "text": str(material.get("text") or ""),
                        "msg": (
                            "Mentinut in SAP"
                            if material_status == "committed"
                            else "Lipseste din raportul SAP final"
                        )
                        if allow_commit
                        else "Dry-run (fara F8 final)",
                    }
                )
            cb(
                {
                    "type": "group-end",
                    "index": idx,
                    "ok": group_ok,
                    "hs_code": hs_code,
                    "status": status,
                    "selected": selected,
                    "requested": len(products),
                    "confirmed": len(committed_products) if allow_commit else 0,
                    "missing": (
                        len(product_rows) - len(committed_products)
                        if allow_commit
                        else 0
                    ),
                    "msg": (
                        f"{len(product_rows) - len(committed_products)} materiale "
                        "lipsesc din confirmarea SAP"
                        if allow_commit and not group_ok
                        else ""
                    ),
                }
            )
        except Exception as e:
            safety_abort = isinstance(e, HsSafetyError)
            summary["failed"] += 1
            summary["materials_failed"] += len(product_rows)
            msg = str(e).splitlines()[0][:200]
            for material in product_rows:
                product = str(material.get("product") or "").strip()
                if not product:
                    continue
                summary["failed_products"].append(product)
                cb(
                    {
                        "type": "material-end",
                        "status": "failed",
                        "scheme": scheme,
                        "product": product,
                        "material": product,
                        "hs_code": hs_code,
                        "text": str(material.get("text") or ""),
                        "msg": msg,
                    }
                )
            cb(
                {
                    "type": "group-end",
                    "index": idx,
                    "ok": False,
                    "hs_code": hs_code,
                    "status": "failed",
                    "msg": msg,
                    "materials": products,
                }
            )
            # Leave SAP on the selection screen so the next group can start.
            try:
                _close_dialog(page, gf)
                gf = _return_to_selection(page, gf, log, timeout_s=40)
            except Exception as recovery_exc:
                log(
                    "warn",
                    f"[{hs_code}] recuperare esuata: {str(recovery_exc)[:100]}",
                )
            if safety_abort:
                summary["safety_aborted"] = True
                summary["abort_reason"] = msg
                summary["skipped"] = len(groups) - idx
                log(
                    "error",
                    "Oprire de siguranta: integritatea selectiei nu mai poate fi "
                    f"garantata. Grupuri ramase: {summary['skipped']}.",
                )
                break
    summary["materials_total_reported"] = (
        summary["materials_committed"]
        + summary["materials_dry_run"]
        + summary["materials_failed"]
        + summary["materials_blocked"]
    )
    return summary


def _classification_completion_error(
    summary: dict, groups: list[dict], allow_commit: bool
) -> str:
    """Describe incomplete HS confirmation. Descriptions still run."""
    expected_materials = sum(
        len([row for row in group.get("products") or [] if row.get("product")])
        for group in groups
    )
    if not allow_commit:
        return "Clasificarea este dry-run; codurile HS nu au fost salvate in SAP."
    if summary.get("safety_aborted"):
        return summary.get("abort_reason") or "Clasificarea HS a fost oprita."
    if summary.get("failed") or summary.get("skipped"):
        return (
            "Clasificarea HS nu este completa: "
            f"grupuri cu eroare={summary.get('failed', 0)}, "
            f"sarite={summary.get('skipped', 0)}."
        )
    if summary.get("committed") != len(groups):
        return (
            "SAP nu a confirmat toate grupurile HS: "
            f"confirmate={summary.get('committed', 0)}/{len(groups)}."
        )
    if summary.get("materials_committed") != expected_materials:
        return (
            "SAP nu a confirmat toate materialele HS: "
            f"confirmate={summary.get('materials_committed', 0)}/"
            f"{expected_materials}."
        )
    return ""


def _navigate_to_sap_app(
    page: Page,
    url: str,
    frame_finder: Callable[..., Frame],
    credentials: dict | None,
    log: Callable[[str, str], None],
    label: str,
    timeout_s: float = 180.0,
    prefer_full_load: bool = False,
) -> Frame:
    """Navigate between the two Fiori WebGUI apps in the same browser."""
    current = page.url or ""
    log("info", f"Navighez la aplicatia SAP {label}")
    log("info", f"URL curent: {_safe_page_url(current)}")
    log("info", f"URL tinta: {_safe_page_url(url)}")
    current_base = current.split("#", 1)[0]
    target_base = str(url or "").split("#", 1)[0]
    current_hash = current.split("#", 1)[1] if "#" in current else ""
    target_hash = str(url or "").split("#", 1)[1] if "#" in str(url or "") else ""
    try:
        if prefer_full_load:
            # A Fiori hash hop leaves the shell without an ITS WebGUI iframe,
            # so this app is opened exactly like the standalone tool does.
            log("info", "Incarc deep link-ul complet, ca in aplicatia standalone.")
            page.goto(url, wait_until="commit", timeout=30000)
        elif current_base and current_base == target_base and target_hash:
            if current_hash != target_hash:
                log("info", "Schimb aplicatia SAP prin hash Fiori, fara reincarcare.")
                page.evaluate(
                    """(hash) => {
                        const next = String(hash || '');
                        if ((location.hash || '') === next) {
                            window.dispatchEvent(new HashChangeEvent('hashchange'));
                        } else {
                            location.hash = next;
                        }
                    }""",
                    "#" + target_hash,
                )
            else:
                log("info", "Hash-ul Fiori este deja pe aplicatia tinta.")
        else:
            page.goto(url, wait_until="commit", timeout=15000)
        log("info", f"Navigare trimisa; URL dupa hop: {_safe_page_url(page.url)}")
    except PWTimeout:
        log("warn", f"Navigarea la {label} a depasit timeout-ul; astept cadrul SAP.")
    except Exception as exc:
        log("warn", f"Navigarea la {label} a esuat: {str(exc)[:160]}")
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=15000)
            log(
                "info",
                f"Reincercare navigare trimisa; URL: {_safe_page_url(page.url)}",
            )
        except Exception as retry_exc:
            log("warn", f"Reincercarea navigarii a esuat: {str(retry_exc)[:160]}")

    deadline = time.time() + timeout_s
    last_progress = 0.0
    last_finder_err = ""
    used_reload = False
    used_full_goto = False
    while time.time() < deadline:
        try:
            frame = _find_frame_any_page(page, frame_finder, 2)
            log(
                "info",
                f"Aplicatia SAP {label} este gata "
                f"({_frame_debug_label(frame.page, frame)})",
            )
            return frame
        except TimeoutError as exc:
            last_finder_err = str(exc).splitlines()[0][:220]
        except Exception as copilot_exc:
            log("warn", f"Cautarea cadrului {label} a esuat: {str(copilot_exc)[:120]}")
        elapsed = timeout_s - max(0.0, deadline - time.time())
        empty_webgui = "probe=product=0" in last_finder_err
        if empty_webgui and not used_reload and elapsed >= 8:
            used_reload = True
            log(
                "warn",
                "Hash-ul Fiori nu a incarcat WebGUI-ul; reincarc pagina "
                "ca in aplicatia standalone de descrieri.",
            )
            try:
                page.reload(wait_until="commit", timeout=20000)
                log("info", f"Reload trimis; URL: {_safe_page_url(page.url)}")
            except Exception as reload_exc:
                log("warn", f"Reload-ul a esuat: {str(reload_exc)[:160]}")
        elif empty_webgui and used_reload and not used_full_goto and elapsed >= 16:
            used_full_goto = True
            log(
                "warn",
                "WebGUI-ul de descrieri lipseste dupa reload; "
                "navighez deep link-ul ca in aplicatia standalone.",
            )
            try:
                page.goto(url, wait_until="commit", timeout=20000)
                log(
                    "info",
                    f"Deep link reincarcata; URL: {_safe_page_url(page.url)}",
                )
            except Exception as goto_exc:
                log("warn", f"Reincarcarea deep link a esuat: {str(goto_exc)[:160]}")
        if elapsed - last_progress >= 8:
            last_progress = elapsed
            log(
                "info",
                f"Astept {label}: {int(elapsed)}s, "
                f"pagini={len(_context_pages(page))}, "
                f"cadre={_frames_debug(page)}, URL={_safe_page_url(page.url)}"
                + (f"; {last_finder_err}" if last_finder_err else ""),
            )
        if _try_form_login(page, credentials, log):
            page.wait_for_timeout(3500)
            continue
        page.wait_for_timeout(400)
    raise TimeoutError(
        f"Aplicatia SAP {label} nu a devenit disponibila in {int(timeout_s)}s. "
        f"URL final: {_safe_page_url(page.url)}"
        + (f" {last_finder_err}" if last_finder_err else "")
    )


def _recover_description_screen(
    page: Page,
    log: Callable[[str, str], None],
) -> None:
    """Close leftover SAP dialogs so the next material starts on selection."""
    try:
        for _ in range(2):
            page.keyboard.press("Escape")
            page.wait_for_timeout(500)
        try:
            _commercial_description_frame(page, timeout_s=3)
            return
        except Exception:
            pass
        page.keyboard.press("F3")
        page.wait_for_timeout(1500)
        try:
            _commercial_description_frame(page, timeout_s=5)
            log("info", "Recovery: am revenit pe ecranul de selectie al descrierilor.")
        except Exception:
            log("warn", "Recovery: ecranul de selectie inca lipseste dupa Escape/F3.")
    except Exception as exc:
        log("warn", f"Recovery dupa eroare a esuat: {str(exc)[:120]}")


def _group_description_items(
    items: list[dict],
    language_by_code: dict,
    planned_languages: list[str],
) -> tuple[list[dict], list[tuple[int, dict, list[str]]]]:
    """Group worklist items that share TECDOC, languages and description text.

    Products under one TECDOC get the same text, so EU-only, GB-only and
    EU+GB products form separate groups that SAP can maintain in one pass.
    """
    grouped: dict[tuple, dict] = {}
    order: list[tuple] = []
    invalid: list[tuple[int, dict, list[str]]] = []
    for idx, item in enumerate(items, start=1):
        descriptions = item.get("descriptions") or {}
        codes = [
            str(code).upper()
            for code in (item.get("languages") or planned_languages)
            if str(code).strip()
        ]
        lang_entries: list[tuple[str, str, str]] = []
        for code in codes:
            language = language_by_code.get(code) or {}
            text = str(descriptions.get(code) or "").strip()
            country = str(language.get("country_name") or "").strip()
            if text and country:
                lang_entries.append((code, text, country))
        if not lang_entries:
            invalid.append((idx, item, codes))
            continue
        tecdoc = str(item.get("tecdoc") or "").strip()
        key = (tecdoc, tuple(lang_entries))
        group = grouped.get(key)
        if group is None:
            group = {
                "tecdoc": tecdoc,
                "codes": codes,
                "lang_entries": lang_entries,
                "members": [],
            }
            grouped[key] = group
            order.append(key)
        group["members"].append((idx, item))
    return [grouped[key] for key in order], invalid


def _maintain_description_bulk(
    page: Page,
    materials: list[str],
    lang_entries: list[tuple[str, str, str]],
    label: str,
    log: Callable[[str, str], None],
) -> None:
    """Maintain one description text for many products in a single SAP pass."""
    for entries in _entries_by_country(lang_entries):
        country = entries[0][2]
        frame = _commercial_description_frame(page, timeout_s=15)
        log(
            "info",
            f"[{label}] mentin in masa {len(materials)} materiale ({country})",
        )
        frame = _paste_products_via_multiple_selection(
            page,
            frame,
            materials,
            log,
            frame_finder=_commercial_description_frame,
        )
        try:
            frame.get_by_role(
                "button", name=re.compile(r"^Execute", re.IGNORECASE)
            ).first.click(timeout=5000)
        except Exception:
            page.keyboard.press("F8")
        status = _wait_for_description_worklist(page, frame, timeout_s=30)
        if status == "empty":
            raise RuntimeError(
                f"SAP nu a gasit niciun material din grupul {label} pentru descrieri."
            )
        if status != "worklist":
            raise RuntimeError(
                f"Worklistul descrierilor nu a aparut pentru grupul {label}."
            )
        result = _mass_maintain_description_selection(page, frame, label, entries, log)
        if result != "ok":
            raise RuntimeError(f"SAP nu a confirmat salvarea in masa: {result}")


def _maintain_one_description(
    page: Page,
    idx: int,
    material: str,
    item_codes: list[str],
    lang_entries: list[tuple[str, str, str]],
    summary: dict,
    cb: Callable[[dict], None],
    log: Callable[[str, str], None],
) -> None:
    try:
        frame = _commercial_description_frame(page, timeout_s=10)
        result = _process_description_material(
            page,
            frame,
            material,
            lang_entries,
            log,
        )
        if result != "ok":
            status = "not_found" if result == "not_found" else "failed"
            if status == "not_found":
                summary["not_found"] += 1
            else:
                summary["failed"] += 1
            cb(
                {
                    "type": "description-item-end",
                    "index": idx,
                    "ok": False,
                    "status": status,
                    "material": material,
                    "languages": item_codes,
                    "msg": f"status SAP neconfirmat: {result}",
                }
            )
            log(
                "warn",
                f"[{material}] descriere comerciala {status}; continui restul worklistului.",
            )
            if status != "not_found":
                _recover_description_screen(page, log)
            return
        summary["saved"] += 1
        saved_codes = [entry[0] for entry in lang_entries]
        cb(
            {
                "type": "description-item-end",
                "index": idx,
                "ok": True,
                "status": "saved",
                "material": material,
                "languages": saved_codes,
                "msg": (
                    "Descrieri "
                    + " si ".join(saved_codes)
                    + " salvate si confirmate de SAP"
                ),
            }
        )
    except Exception as exc:
        summary["failed"] += 1
        msg = str(exc).splitlines()[0][:240]
        cb(
            {
                "type": "description-item-end",
                "index": idx,
                "ok": False,
                "status": "failed",
                "material": material,
                "languages": item_codes,
                "msg": msg,
            }
        )
        log(
            "warn",
            f"[{material}] descrierile nu au fost confirmate: {msg}. "
            "Continui restul materialelor din worklist.",
        )
        _recover_description_screen(page, log)


def _maintain_description_plan(
    page: Page,
    plan: dict,
    languages: list[dict],
    description_url: str,
    classification_url: str,
    credentials: dict | None,
    cancel: threading.Event,
    cb: Callable[[dict], None],
    log: Callable[[str, str], None],
    display_maintained: bool = True,
) -> tuple[dict, Frame]:
    """Save worklist commercial descriptions, then return to Classify Products."""
    items = plan.get("items") or []
    language_by_code = {
        str(language.get("code") or "").strip().upper(): language
        for language in languages
        if language.get("code")
    }
    planned_languages = [
        str(code).upper() for code in (plan.get("languages") or []) if str(code).strip()
    ]
    missing = plan.get("missing") or []
    if missing:
        log(
            "warn",
            f"Sar {len(missing)} materiale fara text DE/EN in referinta; "
            f"continui cu {len(items)} gata din worklist.",
        )

    summary = {
        "total": len(items),
        "saved": 0,
        "failed": 0,
        "not_found": 0,
        "skipped": 0,
        "languages": planned_languages,
    }
    if not items:
        cb({"type": "description-end", "ok": True, "summary": summary})
        return summary, page

    description_frame = _navigate_to_sap_app(
        page,
        description_url,
        _commercial_description_frame,
        credentials,
        log,
        "Manage Customs Commercial Descriptions",
        prefer_full_load=True,
    )
    description_page = description_frame.page
    if description_page is not page:
        log("info", "Tranzactia de descrieri ruleaza intr-un tab SAP separat.")
        try:
            description_page.bring_to_front()
        except Exception:
            pass
    try:
        _set_description_display_maintained(description_frame, display_maintained, log)
    except Exception as exc:
        log("warn", f"Nu am putut activa afisarea produselor mentinute: {exc}")

    cb(
        {
            "type": "description-start",
            "total": len(items),
            "languages": planned_languages,
        }
    )

    groups, invalid = _group_description_items(
        items, language_by_code, planned_languages
    )
    for idx, item, item_codes in invalid:
        material = str(item.get("material") or "").strip()
        summary["failed"] += 1
        log(
            "warn",
            f"[{material}] sar descrierea comerciala: text sau tara SAP lipsesc.",
        )
        cb(
            {
                "type": "description-item-end",
                "index": idx,
                "ok": False,
                "status": "failed",
                "material": material,
                "languages": item_codes,
                "msg": "Text sau tara SAP lipsesc",
            }
        )

    processed = len(invalid)
    for group in groups:
        if cancel.is_set():
            summary["skipped"] = len(items) - processed
            cb({"type": "description-end", "ok": False, "summary": summary})
            raise HsSafetyError(
                "Mentinerea descrierilor a fost anulata; schema urmatoare ramane blocata."
            )
        members = group["members"]
        lang_entries = group["lang_entries"]
        item_codes = group["codes"]
        processed += len(members)
        for idx, item in members:
            cb(
                {
                    "type": "description-item-start",
                    "index": idx,
                    "total": len(items),
                    "material": str(item.get("material") or "").strip(),
                    "tecdoc": item.get("tecdoc", ""),
                    "languages": item_codes,
                }
            )

        materials = [str(item.get("material") or "").strip() for _, item in members]
        bulk_done = False
        if len(materials) > 1:
            label = f"TECDOC {group['tecdoc'] or '-'} x{len(materials)}"
            try:
                _maintain_description_bulk(
                    description_page,
                    materials,
                    lang_entries,
                    label,
                    log,
                )
                bulk_done = True
            except Exception as exc:
                msg = str(exc).splitlines()[0][:240]
                log(
                    "warn",
                    f"[{label}] mentinerea in masa a esuat: {msg}. "
                    "Reiau materialele individual.",
                )
                _recover_description_screen(description_page, log)

        if bulk_done:
            summary["saved"] += len(members)
            saved_codes = [entry[0] for entry in lang_entries]
            for idx, item in members:
                cb(
                    {
                        "type": "description-item-end",
                        "index": idx,
                        "ok": True,
                        "status": "saved",
                        "material": str(item.get("material") or "").strip(),
                        "languages": saved_codes,
                        "msg": (
                            "Descrieri "
                            + " si ".join(saved_codes)
                            + " salvate in masa si confirmate de SAP"
                        ),
                    }
                )
            continue

        for idx, item in members:
            material = str(item.get("material") or "").strip()
            _maintain_one_description(
                description_page,
                idx,
                material,
                item_codes,
                lang_entries,
                summary,
                cb,
                log,
            )

    warning = bool(summary["failed"] or summary["not_found"] or summary["skipped"])
    cb(
        {
            "type": "description-end",
            "ok": True,
            "warning": warning,
            "summary": summary,
        }
    )
    classify_frame = _navigate_to_sap_app(
        page,
        classification_url,
        _classify_frame,
        credentials,
        log,
        "Classify Products",
    )
    return summary, classify_frame


# ============================================================================
# Session
# ============================================================================


class HsSession:
    """Persistent Chromium session for the Classify Products transaction."""

    def __init__(self):
        self._cmd_q: queue.Queue = queue.Queue()
        self._thread: threading.Thread | None = None
        self._lock = threading.Lock()
        self._state = "stopped"
        self._state_msg = ""
        self._cancel = threading.Event()
        self._credentials: dict | None = None

    @property
    def state(self) -> dict:
        # A closed browser leaves the worker thread dead while the state still
        # says ready/busy; report it as stopped so callers restart the session.
        if self._state in ("ready", "busy") and (
            self._thread is None or not self._thread.is_alive()
        ):
            self._state = "stopped"
            self._state_msg = (
                "Sesiunea SAP s-a inchis. Pornesc din nou la urmatorul job."
            )
        return {"state": self._state, "msg": self._state_msg}

    def start(
        self,
        config: dict,
        credentials: dict | None,
        event_cb: Callable[[dict], None],
    ) -> dict:
        with self._lock:
            if self._state in ("ready", "busy", "starting", "stopping"):
                return {"ok": True, "state": self._state}
            self._state = "starting"
            self._state_msg = "Pornesc browserul..."
        self._thread = threading.Thread(
            target=self._worker_loop,
            args=(config, credentials, event_cb),
            daemon=True,
        )
        self._thread.start()
        return {"ok": True, "state": self._state}

    def stop(self) -> None:
        with self._lock:
            if self._state == "stopped":
                return
            self._state = "stopping"
        self._cancel.set()
        self._cmd_q.put(("stop", None))
        thread = self._thread
        if thread:
            thread.join(timeout=15)
        if not thread or not thread.is_alive():
            with self._lock:
                if self._state == "stopping":
                    self._state = "stopped"
                    self._state_msg = "Browser inchis."
            self._thread = None
        else:
            with self._lock:
                self._state_msg = (
                    "Workerul SAP HS nu s-a inchis in 15s; oprirea nu este confirmata."
                )

    def _fail_pending_commands(self, error: str) -> None:
        pending: list[tuple[str, dict]] = []
        with self._lock:
            while True:
                try:
                    command, payload = self._cmd_q.get_nowait()
                except queue.Empty:
                    break
                if command != "stop" and isinstance(payload, dict):
                    pending.append((command, payload))

        for _command, payload in pending:
            callback = payload.get("event_cb")
            if not callable(callback):
                continue
            try:
                callback({"type": "fatal", "error": error})
                callback({"type": "_end_"})
            except Exception:
                pass

    def submit(
        self,
        command: str,
        payload: dict,
        event_cb: Callable[[dict], None],
        cancel_event: threading.Event,
    ) -> None:
        with self._lock:
            current = self._state
            if current in ("ready", "busy") and (
                self._thread is None or not self._thread.is_alive()
            ):
                self._state = "stopped"
                self._state_msg = (
                    "Sesiunea SAP s-a inchis. Pornesc din nou la urmatorul job."
                )
                current = "stopped"
            state_msg = self._state_msg
            if current not in ("error", "stopping", "stopped"):
                self._cancel = cancel_event
                payload = dict(payload)
                payload["event_cb"] = event_cb
                self._cmd_q.put((command, payload))
                return

        event_cb(
            {
                "type": "fatal",
                "error": f"Sesiunea SAP nu este disponibila (state={current}). "
                f"{state_msg}",
            }
        )
        event_cb({"type": "_end_"})

    # -------- worker --------

    def _worker_loop(
        self,
        config: dict,
        credentials: dict | None,
        event_cb: Callable[[dict], None],
    ) -> None:
        self._credentials = credentials
        terminal_error: str | None = None

        def log(level: str, msg: str) -> None:
            event_cb({"type": "log", "level": level, "msg": msg})

        sap_url = config["sap_url"]
        try:
            with sync_playwright() as p:
                if _user_data_locked():
                    raise BrowserBusyError(
                        "Profilul Chromium HS este blocat. Apasa 'Reset browser'."
                    )
                origin = _origin_of(sap_url)
                args = [
                    "--auth-server-allowlist=*.dc.hella.com,*.hella.com",
                    "--auth-negotiate-delegate-allowlist=*.dc.hella.com,*.hella.com",
                    "--ignore-certificate-errors",
                    "--disable-features=IsolateOrigins,site-per-process",
                    '--auto-select-certificate-for-urls=[{"pattern":"'
                    + origin
                    + '","filter":{}}]',
                ]
                http_credentials = None
                if credentials and credentials.get("username"):
                    http_credentials = {
                        "username": str(credentials["username"]),
                        "password": str(credentials.get("password") or ""),
                        "origin": origin,
                    }

                user_data = _user_data_dir()
                user_data.mkdir(parents=True, exist_ok=True)
                _downloads_dir().mkdir(parents=True, exist_ok=True)

                headless = bool(config.get("headless", False))
                log("info", f"Lansez Chromium ({'ascuns' if headless else 'vizibil'})")
                ctx = p.chromium.launch_persistent_context(
                    user_data_dir=str(user_data),
                    headless=headless,
                    args=args,
                    viewport={"width": 1500, "height": 950},
                    accept_downloads=True,
                    ignore_https_errors=True,
                    http_credentials=http_credentials,
                )
                try:
                    ctx.grant_permissions(
                        ["clipboard-read", "clipboard-write"],
                        origin=origin,
                    )
                except Exception:
                    # Some Chromium/SAP combinations don't expose this API for
                    # persistent contexts; fallback keeps using OS clipboard.
                    pass
                page = ctx.pages[0] if ctx.pages else ctx.new_page()

                log("info", "Navighez la tranzactia Classify Products")
                try:
                    page.goto(sap_url, wait_until="commit", timeout=30000)
                except PWTimeout:
                    log("warn", "goto > 30s; astept iframe SAP...")

                gf: Frame | None = None
                deadline = time.time() + 240
                while time.time() < deadline:
                    try:
                        gf = _classify_frame(page, timeout_s=2)
                        break
                    except TimeoutError:
                        pass
                    if _try_form_login(page, credentials, log):
                        page.wait_for_timeout(4000)
                        continue
                    page.wait_for_timeout(2500)

                if gf is None:
                    with self._lock:
                        self._state = "error"
                        self._state_msg = (
                            "Ecranul Classify Products nu a aparut in 4 min."
                        )
                        terminal_error = self._state_msg
                    log("error", self._state_msg)
                    return

                self._state = "ready"
                self._state_msg = "Sesiune SAP activa (Classify Products)."
                try:
                    startup_probe = _frame_probe(gf)
                    log(
                        "info",
                        "Sesiune SAP gata; frame confirmat "
                        f"({_frame_debug_label(page, gf)}; "
                        f"selection={_frame_is_ready(startup_probe)})",
                    )
                except Exception:
                    log("info", "Sesiune SAP gata; frame confirmat.")
                event_cb({"type": "session-ready"})

                while True:
                    try:
                        cmd, payload = self._cmd_q.get(timeout=1.0)
                    except queue.Empty:
                        if not ctx.pages:
                            terminal_error = "Browserul SAP HS s-a inchis inainte de finalizarea comenzilor."
                            break
                        continue
                    if cmd == "stop":
                        log("info", "Inchid sesiunea SAP.")
                        break

                    cb = payload["event_cb"]
                    cb(
                        {
                            "type": "log",
                            "level": "info",
                            "msg": f"Faza command-dequeued: {cmd}",
                        }
                    )
                    self._state = "busy"
                    session_lost = False
                    try:
                        cached_probe = {}
                        if gf is not None:
                            try:
                                cached_probe = _frame_probe(gf)
                            except Exception:
                                cached_probe = {}
                        if gf is None or not _frame_is_ready(cached_probe):
                            cb(
                                {
                                    "type": "log",
                                    "level": "info",
                                    "msg": "Faza frame-reacquire: frameul initial nu mai este valid.",
                                }
                            )
                            gf = _classify_frame(page, timeout_s=30)
                            cached_probe = _frame_probe(gf)
                        cb(
                            {
                                "type": "log",
                                "level": "info",
                                "msg": "Faza command-started: "
                                f"{cmd}; frame {_frame_debug_label(page, gf)}; "
                                f"ready={_frame_is_ready(cached_probe)}",
                            }
                        )
                        if cmd == "download":
                            self._do_download(page, gf, payload, cb)
                        elif cmd == "classify":
                            self._do_classify(page, gf, payload, cb)
                        elif cmd == "descriptions":
                            self._do_descriptions(page, gf, payload, cb)
                        elif cmd == "queue":
                            self._do_queue(page, gf, payload, cb)
                    except Exception as e:
                        try:
                            session_lost = page.is_closed() or not ctx.pages
                        except Exception:
                            session_lost = True
                        cb(
                            {
                                "type": "fatal",
                                "error": f"{e}\n{traceback.format_exc()}",
                            }
                        )
                    finally:
                        cb({"type": "_end_"})
                        self._state = "ready"
                        self._state_msg = "Sesiune SAP gata."

                    if session_lost:
                        log("error", "Browserul SAP s-a inchis; opresc sesiunea.")
                        break

                try:
                    ctx.close()
                except Exception:
                    pass
        except Exception as e:
            error_text = str(e).splitlines()[0][:300] or type(e).__name__
            with self._lock:
                self._state = "error"
                self._state_msg = error_text
                terminal_error = f"Worker HS a esuat: {error_text}"
            event_cb(
                {"type": "log", "level": "error", "msg": f"Worker HS a esuat: {e}"}
            )
            event_cb({"type": "log", "level": "error", "msg": traceback.format_exc()})
        finally:
            self._fail_pending_commands(
                terminal_error
                or "Sesiunea SAP HS s-a inchis inainte de executarea comenzilor."
            )
            with self._lock:
                if self._state != "error":
                    self._state = "stopped"
                    self._state_msg = "Browser inchis."

    def _do_download(self, page: Page, gf: Frame, payload: dict, cb) -> None:
        def log(level: str, msg: str) -> None:
            cb({"type": "log", "level": level, "msg": msg})

        variant = payload["variant"]
        scheme = payload["scheme"]
        display_all = bool(payload.get("display_all_products"))
        cb({"type": "step", "name": "variant", "msg": f"Aplic varianta {variant}"})
        gf = _apply_variant(page, gf, variant, log)

        # The variant may not carry the numbering scheme, and SAP refuses to
        # execute without it.
        cb({"type": "step", "name": "scheme", "msg": f"Setez schema {scheme}"})
        _set_numbering_scheme(page, gf, scheme, log)
        _set_display_all_products(page, gf, display_all, log)

        cb({"type": "step", "name": "execute", "msg": "Execut worklistul"})
        _execute_worklist(page, gf, log)

        cb({"type": "step", "name": "export", "msg": "Descarc worklistul"})
        path = _export_worklist(page, gf, log)
        cb({"type": "download-ready", "path": str(path), "file": Path(path).name})

    def _do_classify(self, page: Page, gf: Frame, payload: dict, cb) -> None:
        def log(level: str, msg: str) -> None:
            cb({"type": "log", "level": level, "msg": msg})

        groups: list[dict] = payload["groups"]
        scheme: str = payload["scheme"]
        variant: str = payload.get("variant") or ""
        blocked: list[dict] = payload.get("blocked") or []
        allow_commit: bool = bool(payload.get("allow_commit"))
        display_all: bool = bool(payload.get("display_all_products"))
        cancel = self._cancel

        if not allow_commit:
            log(
                "warn",
                "MOD DRY-RUN: completez dialogul si ma opresc INAINTE de "
                "'Start Mass Classification' (F8). Nimic nu se salveaza in SAP.",
            )
        log(
            "info",
            f"Varianta SAP folosita per grup: {variant!r}"
            if variant
            else "Fara varianta SAP (doar schema de numerotare).",
        )

        cb({"type": "job-start", "total": len(groups)})
        summary = None
        description_summary = None
        try:
            summary = _classify_groups(
                page,
                gf,
                groups,
                scheme,
                variant,
                allow_commit,
                cancel,
                cb,
                log,
                blocked,
                display_all,
            )
            if allow_commit:
                completion_error = _classification_completion_error(
                    summary, groups, allow_commit
                )
                if completion_error:
                    log(
                        "warn",
                        completion_error + " Continui cu descrierile comerciale pentru "
                        "materialele din worklist.",
                    )
                plan = payload.get("description_plan") or {}
                ready_items = plan.get("items") or []
                log(
                    "info",
                    "Pornesc descrierile comerciale pentru "
                    f"{len(ready_items)} materiale din worklist "
                    f"(HS confirmate={summary.get('materials_committed', 0)}).",
                )
                if ready_items:
                    description_summary, gf = _maintain_description_plan(
                        page,
                        plan,
                        payload.get("description_languages") or [],
                        payload["description_url"],
                        payload["classification_url"],
                        self._credentials,
                        cancel,
                        cb,
                        log,
                        display_maintained=bool(
                            payload.get("description_display_maintained", True)
                        ),
                    )
                else:
                    log(
                        "info",
                        "Nicio descriere comerciala pregatita pentru worklist; "
                        "sar etapa DE/EN.",
                    )
        except Exception as exc:
            if summary is not None:
                summary["description_summary"] = description_summary
            cb(
                {
                    "type": "job-end",
                    "ok": False,
                    "error": str(exc).splitlines()[0][:300],
                    "summary": summary or {},
                }
            )
            raise
        summary["description_summary"] = description_summary
        cb({"type": "job-end", "ok": True, "summary": summary})

    def _do_descriptions(self, page: Page, gf: Frame, payload: dict, cb) -> None:
        """Maintain DE/EN for selected worklist groups without rewriting HS."""

        def log(level: str, msg: str) -> None:
            cb({"type": "log", "level": level, "msg": msg})

        plan = payload.get("description_plan") or {}
        ready_items = plan.get("items") or []
        cancel = self._cancel
        scheme = str(payload.get("scheme") or "").strip()
        cb({"type": "job-start", "total": 0, "scheme": scheme})
        description_summary = None
        try:
            log(
                "info",
                "Pornesc doar descrierile comerciale pentru "
                f"{len(ready_items)} materiale din worklist"
                + (f" ({scheme})." if scheme else "."),
            )
            if not ready_items:
                raise RuntimeError(
                    "Nicio descriere comerciala pregatita pentru grupurile selectate."
                )
            description_summary, gf = _maintain_description_plan(
                page,
                plan,
                payload.get("description_languages") or [],
                payload["description_url"],
                payload["classification_url"],
                self._credentials,
                cancel,
                cb,
                log,
                display_maintained=bool(
                    payload.get("description_display_maintained", True)
                ),
            )
        except Exception as exc:
            cb(
                {
                    "type": "job-end",
                    "ok": False,
                    "error": str(exc).splitlines()[0][:300],
                    "summary": {
                        "description_summary": description_summary,
                        "descriptions_only": True,
                    },
                }
            )
            raise
        cb(
            {
                "type": "job-end",
                "ok": True,
                "summary": {
                    "description_summary": description_summary,
                    "descriptions_only": True,
                    "committed": 0,
                    "failed": int((description_summary or {}).get("failed") or 0),
                    "skipped": int((description_summary or {}).get("skipped") or 0),
                    "materials_committed": 0,
                },
            }
        )

    def _do_queue(self, page: Page, gf: Frame, payload: dict, cb) -> None:
        """Process several (scheme, variant) pairs one after another.

        For each item: apply the variant, set the numbering scheme, export and
        analyze the worklist, then (for a committed run) maintain HS groups and
        commercial descriptions for worklist materials. Incomplete HS confirmation
        does not skip descriptions. A missing SAP variant or empty worklist
        continues with the next scheme. A failed description stage still stops
        the remaining queue.
        """

        def log(level: str, msg: str) -> None:
            cb({"type": "log", "level": level, "msg": msg})

        items: list[dict] = payload["items"]
        auto_classify: bool = bool(payload.get("auto_classify"))
        allow_commit: bool = bool(payload.get("allow_commit"))
        display_all: bool = bool(payload.get("display_all_products"))
        reference_path = payload["reference_path"]
        reference_sheet = payload.get("reference_sheet")
        description_languages: list[dict] = payload.get("description_languages") or []
        description_url: str = payload.get("description_url") or ""
        classification_url: str = payload.get("classification_url") or ""
        display_all_products: bool = bool(payload.get("display_all_products"))
        cancel = self._cancel

        if not allow_commit:
            log(
                "warn",
                "MOD DRY-RUN pentru toata coada: grupurile se completeaza dar "
                "NU se apasa Start Mass Classification (F8).",
            )

        reference = HsReference(reference_path, reference_sheet)
        totals = {
            "schemes": len(items),
            "completed": 0,
            "no_data": 0,
            "no_variant": 0,
            "failed": 0,
            "skipped": 0,
            "safety_aborted": False,
            "abort_reason": "",
        }
        cb({"type": "queue-start", "total": len(items)})

        for idx, item in enumerate(items, start=1):
            if cancel.is_set():
                totals["skipped"] = len(items) - idx + 1
                log("warn", f"Coada anulata de utilizator. Ramase: {totals['skipped']}")
                break

            scheme = item["scheme"]
            variant = item["variant"]
            cb(
                {
                    "type": "scheme-start",
                    "index": idx,
                    "total": len(items),
                    "scheme": scheme,
                    "variant": variant,
                }
            )

            path = None
            result = None
            classify_summary = None
            description_summary = None
            try:
                if idx > 1:
                    gf = _return_to_selection(page, gf, log)

                gf = _apply_variant(page, gf, variant, log)
                _set_numbering_scheme(page, gf, scheme, log)
                _set_display_all_products(page, gf, display_all_products, log)
                _execute_worklist(page, gf, log)
                path = _export_worklist(page, gf, log)

                worklist = read_worklist(path)
                result = analyze(worklist, reference, scheme)
                cb(
                    {
                        "type": "scheme-analysis",
                        "index": idx,
                        "scheme": scheme,
                        "file": Path(path).name,
                        "path": str(path),
                        "result": result,
                    }
                )

                if auto_classify:
                    if result["groups"]:
                        description_plan = None
                        if allow_commit:
                            # Preflight happens before the first irreversible
                            # HS write for this scheme.
                            description_plan = build_description_plan(
                                result["groups"], reference, description_languages
                            )
                            cb(
                                {
                                    "type": "description-preflight",
                                    "index": idx,
                                    "scheme": scheme,
                                    **description_plan["summary"],
                                    "languages": description_plan["languages"],
                                }
                            )
                            if description_plan["missing"]:
                                examples = ", ".join(
                                    f"{row['material']} ({'/'.join(row['missing_languages'])})"
                                    for row in description_plan["missing"][:5]
                                )
                                log(
                                    "warn",
                                    "Sar "
                                    f"{len(description_plan['missing'])} materiale fara text "
                                    "DE/EN in referinta"
                                    + (f": {examples}" if examples else "")
                                    + ". Continui HS si descrierile pentru restul worklistului.",
                                )
                        classify_summary = _classify_groups(
                            page,
                            gf,
                            result["groups"],
                            scheme,
                            variant,
                            allow_commit,
                            cancel,
                            cb,
                            log,
                            result.get("blocked") or [],
                            display_all_products,
                        )
                        completion_error = _classification_completion_error(
                            classify_summary, result["groups"], allow_commit
                        )
                        if completion_error and allow_commit:
                            log(
                                "warn",
                                completion_error
                                + " Continui cu descrierile comerciale pentru "
                                "materialele din worklist.",
                            )
                        if allow_commit:
                            plan = description_plan or {}
                            ready_items = plan.get("items") or []
                            log(
                                "info",
                                "Pornesc descrierile comerciale pentru "
                                f"{len(ready_items)} materiale din worklist "
                                f"(HS confirmate="
                                f"{(classify_summary or {}).get('materials_committed', 0)}).",
                            )
                            if ready_items:
                                description_summary, gf = _maintain_description_plan(
                                    page,
                                    plan,
                                    description_languages,
                                    description_url,
                                    classification_url,
                                    self._credentials,
                                    cancel,
                                    cb,
                                    log,
                                    display_maintained=bool(
                                        payload.get(
                                            "description_display_maintained", True
                                        )
                                    ),
                                )
                            else:
                                log(
                                    "info",
                                    "Nicio descriere comerciala pregatita pentru "
                                    "worklist; sar etapa DE/EN.",
                                )
                    else:
                        log(
                            "info",
                            f"[{scheme}] niciun grup pregatit, nimic de mentinut",
                        )

                totals["completed"] += 1
                cb(
                    {
                        "type": "scheme-end",
                        "index": idx,
                        "scheme": scheme,
                        "ok": True,
                        "analysis_summary": result["summary"],
                        "classify_summary": classify_summary,
                        "description_summary": description_summary,
                        "source_file": Path(path).name,
                        "source_path": str(path),
                    }
                )
            except HsEmptyWorklistError as empty_exc:
                totals["no_data"] += 1
                msg = str(empty_exc).splitlines()[0][:300]
                log("warn", f"[{scheme}] fara date in worklist; trec la urmatoarea.")
                cb(
                    {
                        "type": "scheme-end",
                        "index": idx,
                        "scheme": scheme,
                        "ok": True,
                        "status": "no_data",
                        "msg": msg,
                        "analysis_summary": None,
                        "classify_summary": None,
                        "description_summary": None,
                        "source_file": "",
                        "source_path": "",
                    }
                )
                continue
            except HsMissingVariantError as missing_exc:
                totals["no_variant"] += 1
                msg = str(missing_exc).splitlines()[0][:300]
                log(
                    "warn",
                    f"[{scheme}] varianta SAP lipsa; trec la schema urmatoare.",
                )
                cb(
                    {
                        "type": "scheme-end",
                        "index": idx,
                        "scheme": scheme,
                        "ok": True,
                        "status": "no_variant",
                        "msg": msg,
                        "analysis_summary": None,
                        "classify_summary": None,
                        "description_summary": None,
                        "source_file": "",
                        "source_path": "",
                    }
                )
                continue
            except Exception as e:
                safety_abort = isinstance(e, HsSafetyError)
                totals["failed"] += 1
                msg = str(e).splitlines()[0][:300]
                cb(
                    {
                        "type": "scheme-end",
                        "index": idx,
                        "scheme": scheme,
                        "ok": False,
                        "msg": msg,
                        "analysis_summary": result["summary"] if result else None,
                        "classify_summary": classify_summary,
                        "description_summary": description_summary,
                        "source_file": Path(path).name if path else "",
                        "source_path": str(path) if path else "",
                    }
                )
                log("error", f"[{scheme}] eroare: {msg}")
                totals["safety_aborted"] = safety_abort
                totals["abort_reason"] = msg
                totals["skipped"] = len(items) - idx
                log(
                    "error",
                    "Coada a fost oprita deoarece schema curenta nu este "
                    f"completa; scheme ramase: {totals['skipped']}.",
                )
                break

        cb({"type": "queue-end", "summary": totals})


_session = HsSession()


def get_hs_session() -> HsSession:
    return _session
