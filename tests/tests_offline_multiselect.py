"""Offline checks for the Multiple Selection classification flow.

Reproduces "Multiple Selection for Product": the popup is detected by title,
the Single Value field takes focus, and only Shift+F12 then F8 are sent - no
toolbar buttons, which open the wrong window in SAP.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "server"))

import hs_automation
from playwright.sync_api import sync_playwright

PRODUCTS = ["355135851", "355136431", "355136441"]

POPUP = """
<div id="webguiPopupWindow1" role="dialog">
  <div>Multiple Selection for Product</div>
  <div>
    <span>Select Single Values</span><span>Select Ranges</span>
    <span>Exclude Single Values</span><span>Exclude Ranges</span>
  </div>
  <table>
    <tr><td>S...</td><td>Single Value</td></tr>
    <tr><td><input type="text" id="v0" value="__OLD0__"></td></tr>
    <tr><td><input type="text" id="v1" value="__OLD1__"></td></tr>
    <tr><td><input type="text" id="v2"></td></tr>
    <tr><td><input type="text" id="v3"></td></tr>
  </table>
  <div>
    <div id="M1:48::btn[24]" role="button" title="Import from Text File">up</div>
    <div id="M1:48::btn[8]" role="button" title="Copy (F8)">ok</div>
  </div>
</div>
<script>
window.__events = [];
window.__pasteWorks = __PASTE_WORKS__;
document.addEventListener('click', ev => {
  const btn = ev.target.closest('[role=button]');
  if (btn) window.__events.push('CLICK:' + btn.id);
});
document.addEventListener('keydown', ev => {
  if (ev.key === 'F4' && ev.shiftKey) {
    window.__events.push('SHIFT_F4');
    document.querySelectorAll('#webguiPopupWindow1 input').forEach(
      i => { i.value = ''; }
    );
    return;
  }
  if (ev.key === 'F12' && ev.shiftKey) {
    window.__events.push('SHIFT_F12');
    if (!window.__pasteWorks) return;
    window.__clipboard.split('\\n').forEach((v, i) => {
      const f = document.getElementById('v' + i);
      if (f) f.value = v;
    });
    return;
  }
  if (ev.key === 'F8') {
    window.__events.push('F8');
    document.getElementById('webguiPopupWindow1').style.display = 'none';
  }
});
</script>
"""


def build(paste_works: bool, old: list[str]) -> str:
    return (
        POPUP.replace("__PASTE_WORKS__", "true" if paste_works else "false")
        .replace("__OLD0__", old[0] if len(old) > 0 else "")
        .replace("__OLD1__", old[1] if len(old) > 1 else "")
    )


def drive(page, frame, log) -> None:
    """Run the same popup sequence the automation performs."""
    state = frame.evaluate(hs_automation._JS_MULTI_SELECTION_POPUP_STATE)
    if state.get("filled"):
        log("warn", f"old={state.get('filled')}")
        frame.locator('[data-rpa-multi-input="1"]').first.click()
        page.keyboard.press("Shift+F4")
        page.wait_for_timeout(150)
        state = frame.evaluate(hs_automation._JS_MULTI_SELECTION_POPUP_STATE)
        if state.get("filled"):
            raise RuntimeError("valori vechi ramase in popup")

    hs_automation._set_windows_clipboard_text("\r\n".join(PRODUCTS))
    page.evaluate(
        """async text => {
          if (!navigator.clipboard) {
            Object.defineProperty(navigator, 'clipboard', {
              configurable: true,
              value: {
                writeText: async value => { window.__clipboard = value; },
                readText: async () => window.__clipboard || ''
              }
            });
          }
          await navigator.clipboard.writeText(text);
        }""",
        "\r\n".join(PRODUCTS),
    )
    frame.locator('[data-rpa-multi-input="1"]').first.click()
    page.keyboard.press("Shift+F12")
    page.wait_for_timeout(200)

    after = frame.evaluate(hs_automation._JS_MULTI_SELECTION_POPUP_STATE)
    if after.get("open") and not after.get("filled"):
        raise RuntimeError("Shift+F12 nu a incarcat produsele in popup")

    page.keyboard.press("F8")
    page.wait_for_timeout(200)


def case_happy(page) -> bool:
    page.set_content(build(True, []))
    drive(page, page.main_frame, lambda lvl, m: None)

    events = page.evaluate("window.__events")
    closed = not page.main_frame.evaluate(
        hs_automation._JS_MULTI_SELECTION_POPUP_STATE
    ).get("open")
    ok = events == ["SHIFT_F12", "F8"] and closed
    print(f"HAPPY_PATH: events={events} closed={closed} -> {'PASS' if ok else 'FAIL'}")
    return ok


def case_leftovers(page) -> bool:
    page.set_content(build(True, ["999999", "888888"]))
    drive(page, page.main_frame, lambda lvl, m: None)

    events = page.evaluate("window.__events")
    values = page.evaluate(
        "Array.from(document.querySelectorAll('#webguiPopupWindow1 input'))"
        ".map(i => i.value).filter(Boolean)"
    )
    ok = events == ["SHIFT_F4", "SHIFT_F12", "F8"] and values == PRODUCTS
    print(
        f"LEFTOVERS_CLEARED: events={events} values={values} -> "
        f"{'PASS' if ok else 'FAIL'}"
    )
    return ok


def case_paste_fails(page) -> bool:
    page.set_content(build(False, []))
    try:
        drive(page, page.main_frame, lambda lvl, m: None)
    except RuntimeError as exc:
        events = page.evaluate("window.__events")
        # F8 must never be sent when the paste did not load the products.
        ok = "F8" not in events
        print(
            f"PASTE_FAILURE_ABORTS: {exc} events={events} -> {'PASS' if ok else 'FAIL'}"
        )
        return ok
    print("PASTE_FAILURE_ABORTS: nicio eroare ridicata -> FAIL")
    return False


def main() -> int:
    captured: dict = {}
    hs_automation._set_windows_clipboard_text = lambda t: captured.update(payload=t)

    with sync_playwright() as pw:
        # Let Playwright resolve the installed browser. The old hard-coded
        # chrome-win path made this offline DOM regression fail on macOS/Linux
        # even though the application itself remains Windows-targeted.
        browser = pw.chromium.launch(headless=True)
        context = browser.new_context()
        context.grant_permissions(["clipboard-read", "clipboard-write"])
        context.add_init_script(
            """
            Object.defineProperty(navigator, 'clipboard', {
              configurable: true,
              value: {
                writeText: async text => { window.__clipboard = text; },
                readText: async () => window.__clipboard || ''
              }
            });
            """
        )
        page = context.new_page()
        results = [
            case_happy(page),
            case_leftovers(page),
            case_paste_fails(page),
        ]
        browser.close()

    ok = all(results)
    print("RESULT=", "PASS" if ok else "FAIL")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
