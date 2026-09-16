"""
SAP GTS Fiori automation for "Manage Customs Commercial Descriptions".

Architecture:
  A single global SapSession owns the Playwright loop on a dedicated worker
  thread. Both UI "start session" and "run" reuse this single browser:
    - /session/start  -> spawns the worker (browser opens once)
    - /run            -> sends a job to the worker (no relaunch)
    - /session/stop   -> closes the browser

Validated SAP flow (TEST cift6):
  1. Selection screen: type Product, click Execute
  2. Worklist: detect "no data" -> skip; otherwise click "Column for row
     selection" header (selects all)
  3. Click toolbar "Start Mass Maintenance"
  4. Dialog: dblclick country row (e.g. "Germany")
  5. Click first Language cell, type "DE", Tab, type description
  6. Press F8 to commit
  7. Wait for "Data saved successfully"
  8. Press F3 (back to selection) for next material

Production guard: removed (app is production-ready).
"""

from __future__ import annotations

import queue
import re
import subprocess
import threading
import time
import traceback
from collections.abc import Callable, Iterable
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

from playwright.sync_api import Frame, Page, sync_playwright
from playwright.sync_api import TimeoutError as PWTimeout

# ============================================================================
# user-data lock handling
# ============================================================================


class BrowserBusyError(RuntimeError):
    """Raised when Chromium user-data is locked by another instance."""


def _user_data_dir() -> Path:
    return Path(__file__).resolve().parent.parent / "user-data"


_LOCK_FILES = ("SingletonLock", "SingletonCookie", "SingletonSocket")


def _user_data_locked() -> bool:
    d = _user_data_dir()
    return any((d / f).exists() for f in _LOCK_FILES)


def _list_chrome_pids_using(user_data_path: Path) -> list[int]:
    pids: list[int] = []
    try:
        out = subprocess.check_output(
            [
                "wmic",
                "process",
                "where",
                "name='chrome.exe'",
                "get",
                "ProcessId,CommandLine",
                "/format:csv",
            ],
            stderr=subprocess.DEVNULL,
            text=True,
            timeout=15,
        )
    except Exception:
        return pids
    needle = str(user_data_path).lower()
    for line in out.splitlines():
        if needle in line.lower():
            parts = line.strip().split(",")
            if parts and parts[-1].isdigit():
                pids.append(int(parts[-1]))
    return pids


def cleanup_profile_lock(user_data_path: Path, kill_chrome: bool = True) -> dict:
    """Best-effort cleanup of stale Chromium lock files for one browser profile."""
    d = Path(user_data_path)
    info: dict[str, Any] = {"user_data": str(d), "killed_pids": [], "removed": []}
    if not d.exists():
        return info
    if kill_chrome:
        pids = _list_chrome_pids_using(d)
        for pid in pids:
            try:
                subprocess.run(
                    ["taskkill", "/F", "/PID", str(pid)],
                    check=False,
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    timeout=10,
                )
                info["killed_pids"].append(pid)
            except Exception:
                pass
        if pids:
            time.sleep(1.0)
    for f in _LOCK_FILES:
        p = d / f
        if p.exists():
            try:
                p.unlink()
                info["removed"].append(f)
            except Exception:
                pass
    return info


def cleanup_user_data_lock(kill_chrome: bool = True) -> dict:
    """Best-effort cleanup of stale Chromium lock files."""
    return cleanup_profile_lock(_user_data_dir(), kill_chrome)


# ============================================================================
# helpers
# ============================================================================


def _is_production(url: str, prod_hosts: Iterable[str]) -> bool:
    try:
        host = (urlparse(url).hostname or "").lower()
    except Exception:
        return False
    return any(host == h.lower() or host.endswith("." + h.lower()) for h in prod_hosts)


def _origin_of(url: str) -> str:
    try:
        parsed = urlparse(url)
        port = parsed.port or (443 if parsed.scheme == "https" else 80)
        return f"{parsed.scheme}://{parsed.hostname}:{port}"
    except Exception:
        return "https://*"


def _css_escape(value: str) -> str:
    out = []
    for ch in value:
        if ch.isalnum() or ch == "-" or ch == "_":
            out.append(ch)
        else:
            out.append("\\" + ch)
    return "".join(out)


# ============================================================================
# SAP-specific operations (page must be passed in)
# ============================================================================


def _gui_frame(page: Page, timeout_s: float = 30.0) -> Frame:
    """Find the frame whose body contains the SAP GUI selection screen.

    We require the Product textbox or 'Save as Variant' button to be present
    in the frame's DOM. URL-pattern alone is unreliable because the SAP login
    form is served from the same origin and would match.
    """
    end = time.time() + timeout_s
    last_err = ""
    while time.time() < end:
        for f in page.frames:
            try:
                hit = f.evaluate(
                    """() => {
                        const hasProduct = !!document.querySelector(
                            'input[title=\"Product\"], input[aria-label=\"Product\"]'
                        );
                        const txt = (document.body && document.body.innerText) || '';
                        const hasVariant = /Save as Variant|Variante sichern/i.test(txt);
                        return hasProduct || hasVariant;
                    }"""
                )
                if hit:
                    return f
            except Exception as e:
                last_err = str(e)[:120]
                continue
        page.wait_for_timeout(400)
    raise TimeoutError(
        f"Frame-ul SAP GUI (Product textbox) nu a fost gasit. URL pagina: {page.url} (last_err={last_err})"
    )


def _try_form_login(
    page: Page, credentials: dict | None, log: Callable[[str, str], None]
) -> bool:
    """Fill SAP login form if visible. Returns True if a form was filled."""
    if not credentials or not credentials.get("username"):
        return False
    try:
        user_box = page.locator(
            'input#sap-user, input[name="sap-user"], input#USERNAME_FIELD-inner, '
            'input[name="j_username"]'
        ).first
        pass_box = page.locator(
            'input#sap-password, input[name="sap-password"], input#PASSWORD_FIELD-inner, '
            'input[name="j_password"], input[type="password"]'
        ).first
        if user_box.count() == 0 or pass_box.count() == 0:
            return False
        if not user_box.is_visible():
            return False
        user_box.fill(str(credentials["username"]))
        pass_box.fill(str(credentials["password"]))
        submit = page.locator(
            'button#LOGIN_LINK, button[type="submit"], input[type="submit"], '
            'button:has-text("Log On"), button:has-text("Logon"), button:has-text("Sign In")'
        ).first
        if submit.count() and submit.is_visible():
            submit.click()
        else:
            page.keyboard.press("Enter")
        log("info", "Formular SAP login completat automat.")
        return True
    except Exception:
        return False


def _wait_for_selection_screen(gf: Frame, timeout_s: float = 30.0) -> bool:
    """Wait until 'Product' textbox is visible (we're on selection screen)."""
    end = time.time() + timeout_s
    while time.time() < end:
        try:
            ok = gf.evaluate(
                """() => !!document.querySelector('input[title="Product"], input[aria-label="Product"]')"""
            )
            if ok:
                return True
        except Exception:
            return False
        try:
            gf.wait_for_timeout(400)
        except Exception:
            time.sleep(0.4)
    return False


def _set_display_maintained(
    gf: Frame, want: bool, log: Callable[[str, str], None] | None = None
) -> None:
    """Toggle the SAP 'Display Maintained Products' checkbox on the selection
    screen if its current state differs from `want`. Idempotent and silent
    if the checkbox is not present (e.g. when not on the selection screen)."""
    try:
        state = gf.evaluate(
            """() => {
                const cb = document.querySelector(
                    '[role=checkbox][aria-label="Display Maintained Products"]'
                );
                if (!cb) return null;
                return { id: cb.id, checked: cb.getAttribute('aria-checked') === 'true' };
            }"""
        )
    except Exception:
        return
    if not state:
        return
    if bool(state.get("checked")) == bool(want):
        return
    try:
        gf.locator(f"#{_css_escape(state['id'])}").click(timeout=3000)
        if log:
            log(
                "info",
                f"SAP: 'Display Maintained Products' setat la {want}",
            )
    except Exception as exc:
        if log:
            log("warn", f"Nu am putut comuta 'Display Maintained Products': {exc}")


def _execute_search(page: Page, gf: Frame, material: str) -> None:
    # Both the "from" and "to" range fields have title="Product"; we want
    # the FIRST one (lower x coordinate). Use .first + .fill which is atomic
    # (focus, clear, type) and won't drift focus to the sibling.
    boxes = gf.locator('input[title="Product"]')
    boxes.first.wait_for(state="visible", timeout=15000)
    # Pick the leftmost visible one explicitly
    target = gf.evaluate_handle(
        """() => {
            const all = Array.from(document.querySelectorAll('input[title=\"Product\"]'))
                .filter(el => el.offsetParent);
            all.sort((a, b) => a.getBoundingClientRect().left - b.getBoundingClientRect().left);
            return all[0] || null;
        }"""
    )
    el = target.as_element()
    if el is None:
        raise RuntimeError("Nu am gasit campul Product (left).")
    el.fill(material)
    page.wait_for_timeout(150)
    try:
        gf.get_by_role(
            "button", name=re.compile(r"^Execute", re.IGNORECASE)
        ).first.click(timeout=5000)
    except Exception:
        page.keyboard.press("F8")


def _wait_for_worklist_or_empty(page: Page, gf: Frame, timeout_s: float = 15.0) -> str:
    """After Execute, returns 'worklist', 'empty', or 'unknown'."""
    end = time.time() + timeout_s
    while time.time() < end:
        info = gf.evaluate("""() => {
            const text = (document.body.innerText || '');
            const noData = (/No data (was )?found|Keine Daten gefunden|no entries selected/i.test(text)
                && /No data|Keine Daten/i.test(text));
            // Worklist is open when 'Start Mass Maintenance' toolbar button exists
            const massBtn = !!Array.from(document.querySelectorAll('button, [role=button]'))
                .find(b => /Start Mass Maintenance|Massenpflege/i.test(
                    (b.title || '') + ' ' + (b.getAttribute('aria-label') || '') + ' ' + (b.textContent || '')
                ));
            const hasInfoDialog = !!Array.from(document.querySelectorAll('[role=dialog]'))
                .find(d => /Information/i.test(d.textContent || ''));
            return { noData, massBtn, hasInfoDialog };
        }""")
        if info.get("massBtn"):
            return "worklist"
        if info.get("noData"):
            if info.get("hasInfoDialog"):
                try:
                    page.keyboard.press("Enter")
                    page.wait_for_timeout(400)
                except Exception:
                    pass
            return "empty"
        page.wait_for_timeout(400)
    return "unknown"


_COUNTRY_GRID_JS = """() => {
    const grids = Array.from(document.querySelectorAll('[role=grid]'))
        .filter(g => g.offsetParent);
    // SAP cell ids use the grid's data id, which can differ from the id of
    // the element carrying role=grid (e.g. "<id>-content").
    const cellBase = (g) => {
        const cell = g.querySelector('[id*="["]');
        if (cell) {
            const match = cell.id.match(/^(.*)\\[\\d+,\\d+\\]$/);
            if (match) return match[1];
        }
        const stripped = g.id.replace(/-content$/, '');
        if (stripped !== g.id && document.getElementById(stripped)) return stripped;
        return g.id;
    };
    let countryGrid = null;
    let countryListGrid = null;
    for (const g of grids) {
        const headers = Array.from(g.querySelectorAll('[role=columnheader]'))
            .map(h => h.textContent.trim());
        const hasAbbrev = headers.some(h => /^Descriptn$/i.test(h));
        const hasFull = headers.some(h => /^Description$/i.test(h));
        if (hasAbbrev && !hasFull) {
            countryGrid = cellBase(g);
        }
        const aria = (g.getAttribute('aria-label') || '').toLowerCase();
        if (aria === 'ctries' || aria === 'cties') {
            countryListGrid = g.id;
        }
    }
    if (!countryListGrid) {
        for (const g of grids) {
            const firstCells = Array.from(g.querySelectorAll('[role=row]'))
                .map(r => r.querySelector('[role=gridcell],[role=rowheader]'))
                .filter(Boolean)
                .map(c => (c.textContent || '').trim());
            const codes = firstCells.filter(t => /^[A-Z]{2}$/.test(t));
            if (codes.length >= 3) { countryListGrid = g.id; break; }
        }
    }
    return { countryGrid, countryListGrid };
}"""


def _country_grid_ids(gf: Frame) -> dict:
    """Resolve the Mass Maintenance country grids in the SAP dialog."""
    return gf.evaluate(_COUNTRY_GRID_JS)


def _entries_by_country(
    lang_entries: list[tuple[str, str, str]],
) -> list[list[tuple[str, str, str]]]:
    """Split (code, text, country) entries into one batch per country, in order."""
    grouped: dict[str, list[tuple[str, str, str]]] = {}
    order: list[str] = []
    for entry in lang_entries:
        country = entry[2]
        if country not in grouped:
            grouped[country] = []
            order.append(country)
        grouped[country].append(entry)
    return [grouped[country] for country in order]


def _process_one_material(
    page: Page,
    gf: Frame,
    material: str,
    lang_entries: list[tuple[str, str, str]],
    log: Callable[[str, str], None],
) -> str:
    """Process a single material. Each entry is (language_code, description,
    country_name). SAP cannot revisit a material after save, so every country
    gets its own search + Mass Maintenance + save pass.

    Returns: 'ok' | 'not_found' | raises on failure.
    """
    if not lang_entries:
        raise RuntimeError("Nicio limba activata pentru mentinere.")

    result = "not_found"
    for position, entries in enumerate(_entries_by_country(lang_entries)):
        country = entries[0][2]
        # 0. ensure we're on selection screen
        if not _wait_for_selection_screen(gf, timeout_s=10):
            log("warn", f"[{material}] selection screen nu este vizibil, încerc F3")
            page.keyboard.press("F3")
            page.wait_for_timeout(1500)
            if not _wait_for_selection_screen(gf, timeout_s=15):
                raise RuntimeError("Nu pot reveni la selection screen.")

        # 1. fill Product + Execute
        log("info", f"[{material}] caut... ({country})")
        _execute_search(page, gf, material)

        # 2. wait for worklist OR no-data
        status = _wait_for_worklist_or_empty(page, gf, timeout_s=15)
        if status == "empty":
            if position:
                raise RuntimeError(
                    f"SAP nu mai gaseste {material} pentru tara {country} dupa salvare."
                )
            log("warn", f"[{material}] NU a fost gasit in SAP")
            # _wait_for_worklist_or_empty already pressed Enter to dismiss the
            # Information dialog. We're back on the selection screen, ready
            # for the next material. Do NOT press F3 (it exits the app).
            page.wait_for_timeout(500)
            return "not_found"
        if status != "worklist":
            # no clear response - treat as failure but try to recover
            page.keyboard.press("Escape")
            page.wait_for_timeout(800)
            raise RuntimeError("Worklist nu a aparut in 15s dupa Execute.")

        result = _mass_maintain_selection(page, gf, material, entries, log)
        if result != "ok":
            return result
    return result


def _mass_maintain_selection(
    page: Page,
    gf: Frame,
    material: str,
    lang_entries: list[tuple[str, str, str]],
    log: Callable[[str, str], None],
) -> str:
    """Write the descriptions for every row currently in the worklist.

    `material` is only a log label; the selection may hold many products.
    """
    # 3. select all rows (header checkbox or Ctrl+A in grid)
    try:
        # Try the 'select all' columnheader (label varies between SAP versions)
        sel = gf.locator(
            '[role=columnheader][title*="select all" i], '
            '[role=columnheader]:has-text("select all"), '
            '[role=columnheader]:has-text("Column for row selection")'
        ).first
        sel.click(timeout=3000)
    except Exception:
        # Fallback: click first cell, Ctrl+A
        try:
            gf.locator("[role=grid] [role=row]").first.click(timeout=3000)
            page.keyboard.press("Control+a")
        except Exception:
            log("warn", f"[{material}] nu pot selecta toate, continui oricum")
    page.wait_for_timeout(300)

    # 4. Start Mass Maintenance (find by title attribute)
    mm = gf.locator(
        'button[title*="Start Mass Maintenance" i], '
        '[role=button][title*="Start Mass Maintenance" i]'
    ).first
    mm.click(timeout=10000)
    log("info", f"[{material}] Mass Maintenance")
    dialog = gf.locator("[role=dialog]").first
    dialog.wait_for(state="visible", timeout=15000)
    page.wait_for_timeout(600)

    # 5+6. Mass Maintenance dialog has TWO sections side-by-side:
    #   - "Customs Commercial Description - Global"  (grid header "Description")
    #   - "Customs Commercial Description - Country" (country list `Cties`
    #      on the left, language/description grid `Descriptn` on the right)
    # We must write into the COUNTRY grid only. To switch country, simply
    # click the row in the country list (no dblclick, no F3, no sub-dialog).

    # locate the COUNTRY description grid and the country-list grid.
    # Distinguishing rule (live-verified on material 48603706):
    #   - Global section grid header has "Description" (full word, 11 chars)
    #   - Country section grid header has "Descriptn"  (abbreviated, 9 chars)
    # The country list (Cties) is a grid whose aria-label is "Ctries" and
    # whose first column contains 2-letter country codes.
    grid_ids = _country_grid_ids(gf)
    country_grid_id = grid_ids.get("countryGrid")
    country_list_id = grid_ids.get("countryListGrid")
    if not country_grid_id:
        raise RuntimeError(
            "Nu am gasit grila 'Customs Commercial Description - Country'."
        )
    log(
        "info",
        f"[{material}] grila Country={country_grid_id} lista tari={country_list_id}",
    )

    # Group entries by country in declared order.
    by_country: dict[str, list[tuple[str, str]]] = {}
    for code, desc, ctry in lang_entries:
        by_country.setdefault(ctry, []).append((code, desc))
    countries_in_order = list(by_country.keys())
    log("info", f"[{material}] tari de mentinut: {', '.join(countries_in_order)}")

    for country_name in countries_in_order:
        entries = by_country[country_name]
        log(
            "info",
            f"[{material}] -> {country_name} ({len(entries)} limba/i)",
        )

        # Click the country row in the country list. Match by either the
        # full country name OR a 2-letter code if the user typed e.g. "DE".
        if country_list_id:
            target_cell_id = gf.evaluate(
                """({ listId, name }) => {
                    const list = document.getElementById(listId);
                    if (!list) return null;
                    const rows = Array.from(list.querySelectorAll('[role=row]'));
                    const wanted = name.trim().toLowerCase();
                    for (const r of rows) {
                        const cells = Array.from(
                            r.querySelectorAll('[role=gridcell],[role=rowheader]')
                        );
                        const codeTxt = (cells[0]?.textContent || '').trim().toLowerCase();
                        const nameTxt = (cells[1]?.textContent || '').trim().toLowerCase();
                        if (codeTxt === wanted || nameTxt === wanted) {
                            return cells[0]?.id || null;
                        }
                    }
                    return null;
                }""",
                {"listId": country_list_id, "name": country_name},
            )
        else:
            target_cell_id = None

        # IMPORTANT: a single click only highlights the country row; the right
        # grid (M1:46:3) keeps showing the previously active country. A
        # DOUBLE-click on the row activates it (the country indicator
        # M1:46:::11:47 then shows the new code).
        if target_cell_id:
            gf.locator(f"#{_css_escape(target_cell_id)}").dblclick(timeout=5000)
        else:
            # fallback: dblclick any visible row whose text contains the country name
            row = gf.locator("[role=row]", has_text=country_name).first
            row.wait_for(state="visible", timeout=5000)
            row.dblclick()
        page.wait_for_timeout(700)

        # Confirm the active country switched. The indicator input
        # 'M1:46:::11:47' holds the 2-letter code of the active country.
        active = gf.evaluate(
            """() => {
                const inp = document.getElementById('M1:46:::11:47');
                return inp ? (inp.value || '').trim() : null;
            }"""
        )
        log("info", f"[{material}]   tara activa dupa dblclick: {active!r}")

        # SAP loads already-maintained languages when "Display Maintained
        # Products" is checked. Updating those products means overwriting the
        # existing language row, not appending a duplicate language on the
        # first empty row. Duplicate language rows are rejected by SAP and the
        # editor appears to "vanish" after Enter.
        grid_state = gf.evaluate(
            """(gridId) => {
                function readCell(cell) {
                    if (!cell) return '';
                    const inp = cell.querySelector('input');
                    return (inp ? inp.value : cell.textContent || '').trim();
                }
                function languageCode(value) {
                    const match = String(value || '').trim().match(/^([A-Za-z]{2})\b/);
                    return match ? match[1].toUpperCase() : null;
                }
                const rowsByLanguage = {};
                let firstEmptyRow = 1;
                for (let r = 1; r <= 30; r++) {
                    const cell = document.getElementById(gridId + '[' + r + ',1]');
                    if (!cell) {
                        firstEmptyRow = r;
                        break;
                    }
                    const value = readCell(cell);
                    const code = languageCode(value);
                    if (code && !rowsByLanguage[code]) {
                        rowsByLanguage[code] = r;
                    }
                    if (!value) {
                        firstEmptyRow = r;
                        break;
                    }
                    firstEmptyRow = r + 1;
                }
                return { rowsByLanguage, firstEmptyRow };
            }""",
            country_grid_id,
        )
        rows_by_language = {
            str(code).upper(): int(row)
            for code, row in (grid_state.get("rowsByLanguage") or {}).items()
        }
        used_rows = set(rows_by_language.values())
        next_empty_row = int(grid_state.get("firstEmptyRow") or 1)
        log(
            "info",
            f"[{material}]   randuri existente pe limba: {rows_by_language}, primul liber: {next_empty_row}",
        )

        language_display_names = {
            "DE": "German",
            "EN": "English",
        }

        def read_language_display(cell_id: str) -> str:
            return gf.evaluate(
                """(cellId) => {
                    const cell = document.getElementById(cellId);
                    if (!cell) return '';
                    const editor = cell.querySelector('[id$="_c"]');
                    if (editor) {
                        return (editor.value || editor.textContent || '').trim();
                    }
                    return (cell.textContent || '').trim();
                }""",
                cell_id,
            )

        def select_language_from_dropdown(
            cell_id: str, editor_id: str, code: str
        ) -> None:
            # SAP language is a GuiComboBox. Typing "EN" into it corrupts the
            # autocomplete text ("EN NEnglish"). Open the dropdown and select
            # the exact option by its SAP key instead.
            selected = False
            last_state = {}
            for attempt in range(1, 4):
                gf.locator(f"#{_css_escape(cell_id)}").click(timeout=5000)
                page.wait_for_timeout(250)
                try:
                    gf.locator(f"#{_css_escape(editor_id)}").click(
                        timeout=2500,
                        force=True,
                    )
                except Exception:
                    pass
                # Live-verified in SAP ITS WebGUI: the dropdown helper span can
                # be present but its click does not open the list on empty rows.
                # Alt+ArrowDown on the focused combo input reliably opens the
                # listbox containing data-itemkey="DE"/"EN" options.
                page.keyboard.press("Alt+ArrowDown")
                page.wait_for_timeout(600)
                state = gf.evaluate(
                    """(code) => {
                        const visible = el => !!el && el.offsetParent !== null;
                        const options = Array.from(
                            document.querySelectorAll('[role="option"]')
                        ).filter(visible);
                        const option = options.find(el =>
                            el.getAttribute('data-itemkey') === code ||
                            el.getAttribute('data-itemvalue1') === code
                        );
                        if (option) {
                            option.scrollIntoView({ block: 'center' });
                            option.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
                            option.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
                            option.click();
                            return {
                                selected: true,
                                optionText: (option.textContent || '').trim(),
                                optionId: option.id,
                                optionCount: options.length,
                            };
                        }
                        const listboxes = Array.from(
                            document.querySelectorAll('[role="listbox"]')
                        ).filter(visible).map(lb => ({
                            id: lb.id,
                            text: (lb.textContent || '').trim().slice(0, 160),
                        }));
                        return { selected: false, optionCount: options.length, listboxes };
                    }""",
                    code,
                )
                last_state = state or {}
                if last_state.get("selected"):
                    selected = True
                    log(
                        "info",
                        f"[{material}]   optiune limba {code} selectata din dropdown: {last_state}",
                    )
                    break
                log(
                    "warn",
                    f"[{material}]   dropdown limba {code} fara optiune la incercarea {attempt}: {last_state}",
                )
            if not selected:
                raise RuntimeError(
                    f"Nu pot selecta limba {code} din dropdown SAP. Ultima stare: {last_state}"
                )
            page.wait_for_timeout(1000)
            display = read_language_display(cell_id)
            log("info", f"[{material}]   limba selectata {code}: {display!r}")
            expected_display = language_display_names.get(code)
            if expected_display and display != f"{code} {expected_display}":
                raise RuntimeError(
                    f"Limba {code} nu a fost selectata corect in SAP: {display!r}"
                )

        for language_code, description in entries:
            is_existing_language = language_code in rows_by_language
            if is_existing_language:
                row_idx = rows_by_language[language_code]
                log(
                    "info",
                    f"[{material}]   {language_code}: actualizez rand existent {row_idx}",
                )
            else:
                while next_empty_row in used_rows:
                    next_empty_row += 1
                row_idx = next_empty_row
                used_rows.add(row_idx)
                rows_by_language[language_code] = row_idx
                next_empty_row += 1
                log(
                    "info",
                    f"[{material}]   {language_code}: adaug rand nou {row_idx}",
                )
            # SAP cell ID is [ROW,COL]: col 1 = Language, col 2 = Description.
            lang_cell_id = f"{country_grid_id}[{row_idx},1]"
            desc_cell_id = f"{country_grid_id}[{row_idx},2]"
            lang_editor_id = f"{lang_cell_id}_c"
            desc_editor_id = f"{desc_cell_id}_c"

            current_language_display = read_language_display(lang_cell_id)
            expected_language_display = language_display_names.get(language_code)
            must_select_language = not is_existing_language
            if expected_language_display:
                must_select_language = must_select_language or (
                    current_language_display
                    and current_language_display
                    != f"{language_code} {expected_language_display}"
                )

            if must_select_language:
                # ----- LANGUAGE COLUMN -----
                select_language_from_dropdown(
                    lang_cell_id, lang_editor_id, language_code
                )

            # ----- DESCRIPTION COLUMN -----
            # After Tab focus may have jumped to the language editor of the
            # NEXT row (column-major Tab order). Force-focus the description
            # cell on THIS row.
            gf.locator(f"#{_css_escape(desc_cell_id)}").click(timeout=5000)
            page.wait_for_timeout(300)
            desc_editor = gf.locator(f"#{_css_escape(desc_editor_id)}")
            try:
                desc_editor.wait_for(state="visible", timeout=3000)
            except Exception:
                gf.locator(f"#{_css_escape(desc_cell_id)}").click(timeout=3000)
                desc_editor.wait_for(state="visible", timeout=3000)
            # SAP WebGUI clears values that were only assigned through DOM
            # events. The live-verified sequence is: remove readonly, focus,
            # clear, type real keyboard characters, then press the explicit
            # Continue button.
            ok = gf.evaluate(
                """(inputId) => {
                    const inp = document.getElementById(inputId);
                    if (!inp) return false;
                    try { inp.removeAttribute('readonly'); } catch (e) {}
                    try { inp.readOnly = false; } catch (e) {}
                    inp.focus();
                    inp.value = '';
                    inp.setAttribute('value', '');
                    return true;
                }""",
                desc_editor_id,
            )
            if not ok:
                log(
                    "warn",
                    f"[{material}] editor {desc_editor_id} lipsa, fallback keyboard.type",
                )
            page.keyboard.type(description, delay=10)
            page.wait_for_timeout(200)
            # Explicitly press SAP's Continue button. Raw Enter can be routed
            # to the wrong control in ITS WebGUI and may clear the editor.
            try:
                gf.locator(f"#{_css_escape('M0:50::btn[0]')}").click(
                    timeout=5000,
                    force=True,
                )
            except Exception:
                gf.locator('[role="button"][title*="Continue" i]').first.click(
                    timeout=5000,
                    force=True,
                )
            log(
                "info",
                f"[{material}]   {country_name} rand {row_idx} {language_code}: {description!r} (Continue trimis)",
            )

        page.wait_for_timeout(700)

        # Verify what actually got persisted into this country's grid before we
        # potentially navigate away to the next country.
        def normalize_saved_text(value: Any) -> str:
            return re.sub(r"\s+", " ", str(value or "")).strip()

        persisted = gf.evaluate(
            """(gridId) => {
                function readCell(c) {
                    if (!c) return '';
                    const inp = c.querySelector('input');
                    if (inp && inp.value) return inp.value;
                    return (c.textContent || '').trim();
                }
                const out = [];
                for (let r = 1; r <= 10; r++) {
                    const langCell = document.getElementById(gridId + '[' + r + ',1]');
                    const descCell = document.getElementById(gridId + '[' + r + ',2]');
                    if (!langCell) break;
                    const l = readCell(langCell);
                    const d = readCell(descCell);
                    if (!l && !d) break;
                    out.push({ row: r, lang: l, desc: d });
                }
                return out;
            }""",
            country_grid_id,
        )
        persisted_log = [
            {
                "row": row.get("row"),
                "lang": str(row.get("lang") or "")[:20],
                "desc": str(row.get("desc") or "")[:60],
            }
            for row in persisted
        ]
        log(
            "info",
            f"[{material}]   continut grila Country dupa scriere {country_name}: {persisted_log}",
        )

        missing_entries = []
        for expected_code, expected_description in entries:
            expected_name = language_display_names.get(expected_code)
            expected_language = (
                f"{expected_code} {expected_name}" if expected_name else expected_code
            )
            found = any(
                normalize_saved_text(row.get("lang"))
                == normalize_saved_text(expected_language)
                and normalize_saved_text(row.get("desc"))
                == normalize_saved_text(expected_description)
                for row in persisted
            )
            if not found:
                missing_entries.append(f"{expected_language}: {expected_description}")
        if missing_entries:
            raise RuntimeError(
                "Randurile nu au ramas in grila Country dupa Continue: "
                + " | ".join(missing_entries)
            )

        # The green tick in the country list is useful as a visual clue, but it
        # is not reliable/timely in ITS WebGUI. The grid content above is the
        # source of truth. Do not retry when content is already correct, because
        # retry writes the description twice.
        if target_cell_id and "," in target_cell_id:
            tick_cell_id = target_cell_id.rsplit(",", 1)[0] + ",3"
            tick_ok = gf.evaluate(
                """(cellId) => {
                    const c = document.getElementById(cellId);
                    if (!c) return false;
                    const uses = Array.from(c.querySelectorAll('use'));
                    return uses.some(u => {
                        const h = u.getAttribute('xlink:href') || u.getAttribute('href') || '';
                        return /s_s_okay|okay/i.test(h);
                    });
                }""",
                tick_cell_id,
            )
            log(
                "info",
                f"[{material}]   bifa pentru {country_name} ({tick_cell_id}): {'DA' if tick_ok else 'NU'} (continut OK, fara retry)",
            )
    page.keyboard.press("F8")
    deadline = time.time() + 30
    success = False
    last_msg = ""
    continue_clicked = False
    while time.time() < deadline:
        ok = gf.evaluate("""() => {
            const dlg = Array.from(document.querySelectorAll('[role=dialog]'))
                .filter(p => p.offsetParent);
            const alerts = Array.from(document.querySelectorAll('[role=alert],[role=status]'))
                .map(a => (a.textContent || '').trim()).filter(Boolean);
            // SAP GUI bottom status bar
            const statusBar = document.querySelector('#msgarea, #msgarea-itms, #msgpanel');
            const statusTxt = statusBar ? (statusBar.textContent || '').trim() : '';
            const all = alerts.concat([statusTxt]).filter(Boolean);
            const success = all.some(t => /saved successfully|wurde gesichert|erfolgreich gesichert|data saved/i.test(t));
            const errorMsgs = all.filter(t => /\berror\b|\bfehler\b|cannot|could not|invalid/i.test(t));
            // Detect SAP confirmation dialogs with Continue/OK buttons.
            // "Continue Emphasized" is the actual SAP button text (with \xa0).
            let hasContinueDialog = false;
            let continueBtnId = null;
            for (const d of dlg) {
                const btns = Array.from(d.querySelectorAll('button, [role=button]'));
                const cb = btns.find(b => {
                    const txt = (b.textContent || '').replace(/\\s+/g, ' ').trim().toLowerCase();
                    const title = (b.getAttribute('title') || '').toLowerCase();
                    const aria = (b.getAttribute('aria-label') || '').toLowerCase();
                    return txt.startsWith('continue') || txt === 'ok'
                        || title.includes('continue') || aria.includes('continue');
                });
                if (cb) {
                    hasContinueDialog = true;
                    continueBtnId = cb.id || null;
                }
            }
            return { dialogOpen: dlg.length > 0, success, errors: errorMsgs,
                     statusTxt, alerts, hasContinueDialog, continueBtnId };
        }""")
        if ok:
            last_msg = (ok.get("statusTxt") or "").strip()
            if ok.get("success") and not ok.get("dialogOpen"):
                success = True
                break
            if ok.get("errors"):
                raise RuntimeError("SAP a raportat eroare: " + " | ".join(ok["errors"]))
            # Handle SAP confirmation dialogs (Continue/OK) that block save
            if (
                ok.get("dialogOpen")
                and ok.get("hasContinueDialog")
                and not continue_clicked
            ):
                try:
                    btn_id = ok.get("continueBtnId")
                    if btn_id:
                        gf.locator(f"#{_css_escape(btn_id)}").click(timeout=3000)
                    else:
                        # Button has no ID - press Enter to confirm dialog
                        page.keyboard.press("Enter")
                    continue_clicked = True
                    log("info", f"[{material}] SAP dialog 'Continue' apasat automat")
                    page.wait_for_timeout(2000)
                    continue
                except Exception as cont_err:
                    log("warn", f"[{material}] Nu am putut apasa Continue: {cont_err}")
                    # Fallback: try Enter
                    try:
                        page.keyboard.press("Enter")
                        continue_clicked = True
                        page.wait_for_timeout(2000)
                    except Exception:
                        pass
        page.wait_for_timeout(400)
    if not success:
        # Do NOT press Escape here - let the caller handle recovery
        raise RuntimeError(
            f"Fara confirmare 'Data saved successfully' in 30s. Ultim mesaj SAP: {last_msg!r}"
        )
    log("info", f"[{material}] SAP confirma: {last_msg or 'Data saved successfully'}")

    # 8. back to selection screen
    page.keyboard.press("F3")
    page.wait_for_timeout(1000)
    return "ok"


# ============================================================================
# SapSession - persistent browser owned by a worker thread
# ============================================================================


class SapSession:
    """Owns a single Playwright browser. Commands are dispatched to a worker
    thread that calls Playwright's sync API (which requires single-threaded use)."""

    def __init__(self):
        self._cmd_q: queue.Queue = queue.Queue()
        self._thread: threading.Thread | None = None
        self._lock = threading.Lock()
        self._state: str = "stopped"  # stopped|starting|ready|busy|error|stopping
        self._state_msg: str = ""
        self._cancel = threading.Event()
        self._credentials: dict | None = None
        # public event log buffer for /session/start callers
        self._start_events: list[dict] = []

    # -------- public API --------

    @property
    def state(self) -> dict:
        return {"state": self._state, "msg": self._state_msg}

    def is_ready(self) -> bool:
        return self._state == "ready"

    def start(
        self,
        config: dict,
        credentials: dict | None,
        event_cb: Callable[[dict], None],
    ) -> dict:
        with self._lock:
            if self._state in ("ready", "busy", "starting", "stopping"):
                event_cb(
                    {
                        "type": "log",
                        "level": "info",
                        "msg": f"Sesiunea exista deja (state={self._state}).",
                    }
                )
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

    def run_items(
        self,
        items: list[dict],
        config: dict,
        event_cb: Callable[[dict], None],
        cancel_event: threading.Event,
    ) -> None:
        # The run command is queued regardless of state. The worker thread
        # only consumes commands once it reaches the ready state, so a job
        # queued during startup will execute right after login completes.
        with self._lock:
            state = self._state
            state_msg = self._state_msg
            if state not in ("error", "stopping", "stopped"):
                self._cancel = cancel_event
                done_evt = threading.Event()
                self._cmd_q.put(
                    (
                        "run",
                        {
                            "items": items,
                            "config": config,
                            "event_cb": event_cb,
                            "done": done_evt,
                        },
                    )
                )
                # don't block API thread; the worker will signal _end_ via event_cb
                return

        event_cb(
            {
                "type": "fatal",
                "error": f"Sesiunea SAP nu este disponibila (state={state}). "
                f"{state_msg}",
            }
        )
        event_cb({"type": "_end_"})

    def _fail_pending_commands(self, error: str) -> None:
        pending: list[tuple[str, dict]] = []
        with self._lock:
            while True:
                try:
                    cmd, payload = self._cmd_q.get_nowait()
                except queue.Empty:
                    break
                if cmd != "stop" and isinstance(payload, dict):
                    pending.append((cmd, payload))

        for _cmd, payload in pending:
            callback = payload.get("event_cb")
            if not callable(callback):
                continue
            try:
                callback({"type": "fatal", "error": error})
                callback({"type": "_end_"})
            except Exception:
                pass
            finally:
                done = payload.get("done")
                if done is not None:
                    done.set()

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
                    "Workerul SAP nu s-a inchis in 15s; oprirea nu este confirmata."
                )

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
                        "user-data este blocat. Apasa 'Reseteaza sesiunea' mai intai."
                    )
                origin = _origin_of(sap_url)
                auto_cert = (
                    '--auto-select-certificate-for-urls=[{"pattern":"'
                    + origin
                    + '","filter":{}}]'
                )
                args = [
                    "--auth-server-allowlist=*.dc.hella.com,*.hella.com",
                    "--auth-negotiate-delegate-allowlist=*.dc.hella.com,*.hella.com",
                    "--ignore-certificate-errors",
                    "--disable-features=IsolateOrigins,site-per-process",
                    auto_cert,
                ]
                http_credentials = None
                if (
                    credentials
                    and credentials.get("username")
                    and credentials.get("password")
                ):
                    http_credentials = {
                        "username": str(credentials["username"]),
                        "password": str(credentials["password"]),
                        "origin": origin,
                    }

                user_data = _user_data_dir()
                user_data.mkdir(parents=True, exist_ok=True)

                headless = bool(config.get("headless", False))
                mode_label = "ascuns (headless)" if headless else "vizibil"
                log("info", f"Lansez Chromium ({mode_label})...")
                ctx = p.chromium.launch_persistent_context(
                    user_data_dir=str(user_data),
                    headless=headless,
                    args=args,
                    viewport={"width": 1400, "height": 900},
                    accept_downloads=False,
                    ignore_https_errors=True,
                    http_credentials=http_credentials,
                )
                page = ctx.pages[0] if ctx.pages else ctx.new_page()

                log("info", f"Navighez la {sap_url}")
                try:
                    page.goto(sap_url, wait_until="commit", timeout=30000)
                except PWTimeout:
                    log("warn", "goto > 30s; astept iframe SAP...")

                # Continuous login + GUI-detect loop. SAP often shows the
                # logon form multiple times (once for Fiori, once for the
                # embedded GUI app), so we must keep filling it until we
                # actually see the Product textbox.
                gf: Frame | None = None
                deadline = time.time() + 240  # 4 min total
                fills = 0
                last_url = ""
                while time.time() < deadline:
                    # 1. is the GUI ready?
                    try:
                        gf = _gui_frame(page, timeout_s=2)
                        break
                    except TimeoutError:
                        pass
                    # 2. is a login form visible? fill it
                    if _try_form_login(page, credentials, log):
                        fills += 1
                        log("info", f"Formular completat (#{fills}), astept...")
                        page.wait_for_timeout(4000)
                        continue
                    # 3. log progress every ~10s
                    cur = page.url or ""
                    if cur != last_url:
                        log("info", f"URL: {cur[:160]}")
                        last_url = cur
                    page.wait_for_timeout(2500)

                if gf is None:
                    with self._lock:
                        self._state = "error"
                        self._state_msg = "GUI SAP nu a aparut in 4 min."
                        terminal_error = self._state_msg
                    log("error", self._state_msg)
                    log("warn", f"URL final: {page.url}")
                    return

                if not _wait_for_selection_screen(gf, timeout_s=30):
                    log(
                        "warn",
                        "Product textbox nu este detectat - posibil ecran neasteptat.",
                    )

                self._state = "ready"
                self._state_msg = "Sesiune SAP activa, gata pentru job-uri."
                log("info", "Sesiune SAP gata.")
                event_cb({"type": "session-ready"})

                # Command loop
                while True:
                    try:
                        cmd, payload = self._cmd_q.get(timeout=1.0)
                    except queue.Empty:
                        # check that browser context still alive
                        if not ctx.pages:
                            log(
                                "warn",
                                "Toate pagini Chromium inchise; opresc sesiunea.",
                            )
                            terminal_error = "Browserul SAP s-a inchis inainte de finalizarea comenzilor."
                            break
                        continue

                    if cmd == "stop":
                        log("info", "Inchid sesiunea SAP.")
                        break

                    if cmd == "run":
                        self._state = "busy"
                        try:
                            self._run_job(page, gf, payload)
                        except Exception as e:
                            tb = traceback.format_exc()
                            payload["event_cb"](
                                {
                                    "type": "fatal",
                                    "error": f"{e}\n{tb}",
                                }
                            )
                        finally:
                            payload["event_cb"]({"type": "_end_"})
                            payload["done"].set()
                            self._state = "ready"
                            self._state_msg = "Sesiune SAP gata."

                try:
                    ctx.close()
                except Exception:
                    pass
        except Exception as e:
            error_text = str(e).splitlines()[0][:300] or type(e).__name__
            with self._lock:
                self._state = "error"
                self._state_msg = error_text
                terminal_error = f"Worker SAP a esuat: {error_text}"
            event_cb(
                {"type": "log", "level": "error", "msg": f"Worker SAP a esuat: {e}"}
            )
            event_cb({"type": "log", "level": "error", "msg": traceback.format_exc()})
        finally:
            self._fail_pending_commands(
                terminal_error
                or "Sesiunea SAP s-a inchis inainte de executarea comenzilor."
            )
            with self._lock:
                if self._state != "error":
                    self._state = "stopped"
                    self._state_msg = "Browser inchis."

    def _run_job(self, page: Page, gf: Frame, payload: dict) -> None:
        items = payload["items"]
        config = payload["config"]
        event_cb = payload["event_cb"]
        cancel = self._cancel
        default_country = config.get("country_name", "Germany")

        # Active languages from config (in declared order). Each entry has
        # its own SAP country (e.g. DE -> Germany, EN -> United Kingdom).
        cfg_langs = config.get("languages") or [
            {
                "code": (config.get("language") or "DE").upper(),
                "country_name": default_country,
                "enabled": True,
            }
        ]
        active_langs = [
            {
                "code": (lng.get("code") or "").upper(),
                "country_name": (lng.get("country_name") or default_country).strip(),
            }
            for lng in cfg_langs
            if lng.get("enabled") and lng.get("code")
        ]
        active_codes = [lng["code"] for lng in active_langs]
        country_by_code = {lng["code"]: lng["country_name"] for lng in active_langs}

        def log(level: str, msg: str) -> None:
            event_cb({"type": "log", "level": level, "msg": msg})

        if not active_codes:
            log("error", "Nicio limba activata in configurare. Opresc job-ul.")
            event_cb(
                {
                    "type": "job-end",
                    "summary": {
                        "total": len(items),
                        "ok": 0,
                        "fail": 0,
                        "not_found": 0,
                        "skipped": len(items),
                    },
                }
            )
            return

        log("info", f"Limbi mentinute: {', '.join(active_codes)}")

        summary = {
            "total": len(items),
            "ok": 0,
            "fail": 0,
            "not_found": 0,
            "skipped": 0,
        }
        event_cb({"type": "job-start", "total": len(items)})

        # Apply 'Display Maintained Products' preference once per job, on the
        # selection screen. Allows re-maintenance of already-saved products.
        want_display_maintained = bool(config.get("display_maintained", True))
        try:
            gf_init = _gui_frame(page, timeout_s=5)
            if _wait_for_selection_screen(gf_init, timeout_s=5):
                _set_display_maintained(gf_init, want_display_maintained, log)
        except Exception:
            pass

        for idx, item in enumerate(items, start=1):
            if cancel.is_set():
                summary["skipped"] = len(items) - idx + 1
                log("warn", f"Anulat de utilizator. Skipped: {summary['skipped']}")
                break

            material = str(item.get("material", "")).strip()
            descriptions = item.get("descriptions") or {}
            # legacy fallback if frontend sent flat 'description'
            if not descriptions and item.get("description"):
                descriptions = {active_codes[0]: str(item["description"]).strip()}
            # normalize keys to upper
            descriptions = {
                str(k).upper(): str(v).strip()
                for k, v in descriptions.items()
                if v and str(v).strip()
            }
            row = item.get("row")

            # Build lang entries in config-declared order, only for codes that
            # have a non-empty description for this material. Each entry is
            # (code, description, country_name).
            lang_entries: list[tuple[str, str, str]] = [
                (code, descriptions[code], country_by_code[code])
                for code in active_codes
                if code in descriptions
            ]

            event_cb(
                {
                    "type": "item-start",
                    "index": idx,
                    "total": len(items),
                    "row": row,
                    "material": material,
                    "description": " | ".join(
                        f"{c}@{ctry}: {d}" for c, d, ctry in lang_entries
                    ),
                }
            )

            if not material or not lang_entries:
                summary["fail"] += 1
                event_cb(
                    {
                        "type": "item-end",
                        "index": idx,
                        "ok": False,
                        "material": material,
                        "msg": "lipsa material sau descriere pentru limbile active",
                    }
                )
                continue

            try:
                # Refresh frame ref. If we drifted to Shell-home (e.g. F3
                # accidentally exited the app), navigate back to the deep link.
                try:
                    gf = _gui_frame(page, timeout_s=8)
                except TimeoutError:
                    log(
                        "warn",
                        f"GUI nu mai e vizibil; re-navighez la app... (URL: {page.url[:120]})",
                    )
                    sap_url = config["sap_url"]
                    # Extract the app hash identifier from the deep link so we
                    # can detect when SAP redirects us to #Shell-home instead.
                    _app_hash_marker = (
                        sap_url.split("#")[1].split("?")[0] if "#" in sap_url else ""
                    )
                    # Navigate back to the app deep link
                    try:
                        page.goto(sap_url, wait_until="commit", timeout=30000)
                    except Exception:
                        pass
                    # Wait for page to load and handle potential login forms.
                    # Recovery strategy (in order):
                    #  1. If SAP redirects to #Shell-home (wrong URL) -> goto(sap_url), max 3x
                    #  2. If URL is correct but GUI frame still missing after 30s
                    #     -> page.reload() to force SAP ITS to restart the GUI session
                    #  3. If still missing after reload+45s -> goto(sap_url) once more
                    nav_deadline = time.time() + 150
                    last_action_time = time.time()
                    nav_retries = 0
                    reload_done = False
                    gf = None
                    while time.time() < nav_deadline:
                        # Check if GUI frame appeared
                        try:
                            gf = _gui_frame(page, timeout_s=3)
                            break
                        except TimeoutError:
                            pass
                        # Check for login form and fill it
                        if _try_form_login(page, self._credentials, log):
                            log("info", "Re-login SAP dupa recovery...")
                            page.wait_for_timeout(4000)
                            last_action_time = time.time()
                            continue
                        cur_url = page.url or ""
                        on_app = _app_hash_marker and _app_hash_marker in cur_url
                        elapsed = time.time() - last_action_time
                        if not on_app and elapsed > 15 and nav_retries < 3:
                            # Wrong URL (e.g. #Shell-home) - navigate to app
                            nav_retries += 1
                            log(
                                "info",
                                f"SAP pe URL gresit ({cur_url[cur_url.find('#') : cur_url.find('#') + 50] or cur_url[-50:]}); "
                                f"re-navighez la app (tentativa {nav_retries}/3)...",
                            )
                            try:
                                page.goto(sap_url, wait_until="commit", timeout=30000)
                            except Exception:
                                pass
                            last_action_time = time.time()
                            page.wait_for_timeout(5000)
                            continue
                        if on_app and not reload_done and elapsed > 30:
                            # Correct URL but GUI iframe stuck (SAP ITS session broken)
                            # A reload forces SAP to start a fresh GUI session.
                            reload_done = True
                            log(
                                "info",
                                "GUI frame absent desi URL e corect; reîncarc pagina SAP...",
                            )
                            try:
                                page.reload(wait_until="commit", timeout=30000)
                            except Exception:
                                pass
                            last_action_time = time.time()
                            page.wait_for_timeout(5000)
                            continue
                        if on_app and reload_done and elapsed > 45 and nav_retries < 3:
                            # Still broken after reload - force a full goto
                            nav_retries += 1
                            log(
                                "info",
                                f"GUI frame inca absent dupa reload; re-navighez complet (tentativa {nav_retries}/3)...",
                            )
                            try:
                                page.goto(sap_url, wait_until="commit", timeout=30000)
                            except Exception:
                                pass
                            last_action_time = time.time()
                            reload_done = False
                            page.wait_for_timeout(5000)
                            continue
                        page.wait_for_timeout(3000)
                    if gf is None:
                        log(
                            "error",
                            f"EROARE: Frame-ul SAP GUI (Product textbox) nu a fost gasit. "
                            f"URL pagina: {page.url} (last_err=)",
                        )
                        raise TimeoutError(
                            f"Frame-ul SAP GUI (Product textbox) nu a fost gasit. "
                            f"URL pagina: {page.url} (last_err=)"
                        )
                    # After successful re-navigation, re-apply the
                    # 'Display Maintained Products' checkbox setting.
                    try:
                        _set_display_maintained(
                            gf, bool(config.get("display_maintained", True)), log
                        )
                    except Exception:
                        pass
                result = _process_one_material(
                    page,
                    gf,
                    material,
                    lang_entries,
                    log,
                )
                if result == "ok":
                    summary["ok"] += 1
                    event_cb(
                        {
                            "type": "item-end",
                            "index": idx,
                            "ok": True,
                            "material": material,
                            "msg": "Saved",
                        }
                    )
                elif result == "not_found":
                    summary["not_found"] += 1
                    event_cb(
                        {
                            "type": "item-end",
                            "index": idx,
                            "ok": False,
                            "material": material,
                            "msg": "Material negasit in SAP",
                            "not_found": True,
                        }
                    )
                else:
                    summary["fail"] += 1
                    event_cb(
                        {
                            "type": "item-end",
                            "index": idx,
                            "ok": False,
                            "material": material,
                            "msg": f"status: {result}",
                        }
                    )
            except Exception as e:
                summary["fail"] += 1
                err_msg = str(e).splitlines()[0][:200]
                event_cb(
                    {
                        "type": "item-end",
                        "index": idx,
                        "ok": False,
                        "material": material,
                        "msg": err_msg,
                    }
                )
                # Smart recovery: dismiss dialogs but avoid F3 which can exit
                # the app entirely. Instead, navigate back to the app URL.
                try:
                    page.keyboard.press("Escape")
                    page.wait_for_timeout(500)
                    page.keyboard.press("Escape")
                    page.wait_for_timeout(500)
                    # Check if we're still in the app (Product textbox visible)
                    try:
                        gf_check = _gui_frame(page, timeout_s=3)
                        if _wait_for_selection_screen(gf_check, timeout_s=3):
                            log("info", "Recovery: inca pe selection screen, continui")
                        else:
                            # We're in worklist or mass maintenance - F3 back
                            page.keyboard.press("F3")
                            page.wait_for_timeout(1500)
                            # Check again - if F3 brought us to selection screen
                            try:
                                gf_check2 = _gui_frame(page, timeout_s=3)
                                if not _wait_for_selection_screen(
                                    gf_check2, timeout_s=3
                                ):
                                    page.keyboard.press("F3")
                                    page.wait_for_timeout(1500)
                            except TimeoutError:
                                pass
                    except TimeoutError:
                        # GUI frame lost - will be handled by re-navigation
                        # at the top of the next iteration
                        log("warn", "Recovery: GUI frame pierdut, voi re-naviga")
                except Exception:
                    pass

        event_cb({"type": "job-end", "summary": summary})


# ============================================================================
# Module-level singleton
# ============================================================================

_session = SapSession()


def get_session() -> SapSession:
    return _session


def run_update(
    items: list[dict],
    config: dict,
    progress_cb: Callable[[dict], None] | None = None,
    cancel_event: threading.Event | None = None,
    headless: bool | None = None,
    credentials: dict | None = None,
    startup_timeout_s: float = 300.0,
) -> dict:
    """Compatibility helper for older scripts that ran one SAP update batch.

    The web app now uses the persistent SapSession API directly. This wrapper
    keeps local scripts such as smoke_test.py working by starting the singleton
    session when needed, waiting for the job to finish, collecting results, and
    closing only the session it started itself.
    """
    cfg = dict(config)
    if headless is not None:
        cfg["headless"] = bool(headless)

    cancel = cancel_event or threading.Event()
    done = threading.Event()
    result: dict[str, Any] = {"summary": None, "items": [], "fatal": None}

    def on_event(ev: dict) -> None:
        ev_type = ev.get("type")
        if ev_type == "job-end":
            result["summary"] = ev.get("summary")
        elif ev_type == "item-end":
            result["items"].append(
                {
                    "material": ev.get("material"),
                    "ok": ev.get("ok"),
                    "msg": ev.get("msg"),
                    "not_found": ev.get("not_found", False),
                }
            )
        elif ev_type == "fatal":
            result["fatal"] = ev.get("error")
        elif ev_type == "_end_":
            done.set()
        if progress_cb:
            progress_cb(ev)

    sess = get_session()
    initial_state = sess.state["state"]
    if initial_state == "busy":
        raise RuntimeError("Sesiunea SAP este deja ocupata cu alt job.")

    started_here = initial_state in ("stopped", "error")
    try:
        if started_here:
            sess.start(cfg, credentials, on_event)

        deadline = time.time() + startup_timeout_s
        while sess.state["state"] == "starting":
            if cancel.is_set():
                sess.stop()
                raise RuntimeError("Pornirea sesiunii SAP a fost anulata.")
            if time.time() > deadline:
                raise TimeoutError(
                    f"Sesiunea SAP nu a devenit ready in {startup_timeout_s:.0f}s."
                )
            time.sleep(0.25)

        state = sess.state
        if state["state"] == "error":
            raise RuntimeError(state.get("msg") or "Sesiunea SAP este in eroare.")
        if state["state"] == "stopped":
            raise RuntimeError("Sesiunea SAP s-a inchis inainte de rularea jobului.")

        sess.run_items(items, cfg, event_cb=on_event, cancel_event=cancel)
        while not done.wait(0.25):
            state = sess.state
            if state["state"] == "error":
                raise RuntimeError(state.get("msg") or "Jobul SAP a esuat.")
            if state["state"] == "stopped":
                raise RuntimeError(
                    "Sesiunea SAP s-a inchis inainte de finalizarea jobului."
                )

        if result.get("fatal"):
            raise RuntimeError(str(result["fatal"]))
        return {"summary": result["summary"], "items": result["items"]}
    finally:
        if started_here:
            sess.stop()
