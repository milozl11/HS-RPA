"""Offline regression for HS field confirmation and guarded SAP commit.

The synthetic page models the important SAP ITS contract:
one Enter validates the tariff field, while Start Mass Classification (or F8
fallback) opens a report that explicitly lists the classified products.
"""

import sys
import threading
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "server"))

import hs_automation
from playwright.sync_api import sync_playwright


def dialog_html(*, include_start: bool = True, unrelated_popup: bool = False) -> str:
    start = (
        '<div id="C640_toolbar_btn8" role="button" '
        'title="Start Mass Classification (F8)">Start</div>'
        if include_start
        else ""
    )
    unrelated = (
        '<div id="unrelated" role="dialog"><label for="user">User</label>'
        '<input id="user" title="User Name"></div>'
        if unrelated_popup
        else ""
    )
    return f"""
{start}
<div id="webguiPopupWindow2" role="dialog">
  <div>Mass Classification</div>
  <label for="scheme">Numbering Scheme</label>
  <input id="scheme" title="Numbering Scheme" value="EDCHSCDEEX">
  <label for="tariff">Tariff Number</label>
  <input id="tariff" title="Tariff Number" value="">
  <div id="cancel" role="button" title="Cancel (F12)">Cancel</div>
</div>
{unrelated}
<script>
window.__events = [];
window.__openReport = source => {{
  window.__events.push('START:' + source);
  const mass = document.getElementById('webguiPopupWindow2');
  if (mass) mass.style.display = 'none';
  const unrelated = document.getElementById('unrelated');
  if (unrelated) unrelated.style.display = 'none';
  const report = document.createElement('div');
  report.id = 'SAPMSSY_REPORT';
  report.setAttribute('role', 'dialog');
  report.innerHTML = `
    <div id="userarealist1">
      <div class="lsAbapList__item">Product 355135851 is classified with number 90328900</div>
      <div class="lsAbapList__item">Product 355136431 is classified with number 90328900</div>
    </div>
    <div role="button" title="Continue (Enter)" id="continue">Continue</div>`;
  document.body.appendChild(report);
  report.querySelector('#continue').addEventListener('click', () => report.remove());
}};
document.addEventListener('click', ev => {{
  const btn = ev.target.closest('[role=button]');
  if (!btn) return;
  window.__events.push('CLICK:' + (btn.getAttribute('title') || btn.id));
  if ((btn.getAttribute('title') || '').includes('Start Mass Classification')) {{
    window.__openReport('CLICK');
  }}
  if ((btn.getAttribute('title') || '').includes('Cancel')) {{
    const dialog = btn.closest('[role=dialog]');
    if (dialog) dialog.style.display = 'none';
  }}
}});
document.addEventListener('keydown', ev => {{
  window.__events.push('KEY:' + ev.key);
  if (ev.key === 'F8') window.__openReport('F8');
}});
</script>
"""


def run_case(page, *, include_start: bool, unrelated_popup: bool = False) -> None:
    page.set_content(
        dialog_html(include_start=include_start, unrelated_popup=unrelated_popup)
    )
    frame = page.main_frame
    logs: list[tuple[str, str]] = []

    def log(level: str, msg: str) -> None:
        logs.append((level, msg))

    probe = frame.evaluate(hs_automation._JS_MASS_DIALOG_PROBE, "EDCHSCDEEX")
    assert probe.get("open") and probe.get("relevant"), probe
    assert probe.get("dialogId") == "webguiPopupWindow2", probe
    assert probe.get("targetId") == "tariff", probe

    # Filling and validation are not allowed to start the SAP action.
    hs_automation._fill_tariff_code(page, frame, "EDCHSCDEEX", "90328900", log)
    events = page.evaluate("window.__events")
    assert events.count("KEY:Enter") == 1, events
    assert not any(e.startswith("START:") for e in events), events
    assert page.locator("#webguiPopupWindow2").is_visible()

    committed = hs_automation._commit_classification(
        page,
        frame,
        log,
        ["355135851", "355136431"],
    )
    assert committed == {"355135851", "355136431"}, (committed, logs)

    events = page.evaluate("window.__events")
    expected_source = "START:CLICK" if include_start else "START:F8"
    assert expected_source in events, events
    if include_start:
        assert "START:F8" not in events, events
    assert not page.locator("#SAPMSSY_REPORT").count(), events


def test_hidden_dialog_does_not_block_selection(page) -> None:
    page.set_content(
        """
        <div role="button" title="Get Variant">Get Variant</div>
        <div role="dialog" id="old-popup" style="display:none">Old popup</div>
        <script>
          window.__keys = [];
          document.addEventListener('keydown', ev => window.__keys.push(ev.key));
        </script>
        """
    )
    frame = page.main_frame
    returned = hs_automation._return_to_selection(
        page, frame, lambda level, msg: None, timeout_s=0.5
    )
    assert returned == frame
    assert page.evaluate("window.__keys") == []


def test_classifier_prefers_ready_frame_over_stale_frame(page) -> None:
    page.set_content(
        """
        <div title="Get Variant">Stale Get Variant</div>
        <iframe srcdoc=''
            title="sap-frame"></iframe>
        """
    )
    page.frames[1].set_content(
        """
        <div title="Get Variant">Get Variant</div>
        <span id="scheme-label">Numbering Scheme</span>
        <input id="scheme" aria-labelledby="scheme-label" value="EDCHSCDEEX">
        """
    )
    returned = hs_automation._classify_frame(page, timeout_s=1)
    assert returned == page.frames[1]
    assert hs_automation._frame_is_ready(hs_automation._frame_probe(returned))


def test_variant_dialog_can_be_found_in_child_frame(page) -> None:
    page.set_content(
        """
        <div title="Get Variant">Get Variant</div>
        <input title="Numbering Scheme" value="EDCHSCDEEX">
        <iframe srcdoc='<div role="dialog" id="variant-popup">
            <div>Get Variant</div>
            <input title="Variant Name">
            <input title="User Name">
        </div>'></iframe>
        """
    )
    selection = hs_automation._classify_frame(page, timeout_s=1)
    dialog_frame, dialog = hs_automation._find_variant_dialog(
        page, selection, timeout_s=1
    )
    assert dialog_frame == page.frames[1]
    assert dialog.is_visible()


def test_apply_variant_completes_selection_flow(page) -> None:
    page.set_content(
        """
                <div title="Get Variant">Get Variant</div>
                <input title="Numbering Scheme" value="EDCHSCDEEX">
                <script>
                    window.__variant = {};
                    document.querySelector('[title="Get Variant"]').addEventListener('click', () => {
                        const dialog = document.createElement('div');
                        dialog.id = 'variant-dialog';
                        dialog.setAttribute('role', 'dialog');
                        dialog.innerHTML = `
                            <input title="User Name" value="CURRENT">
                            <input title="Variant Name">
                            <button title="Execute">Execute</button>`;
                        document.body.appendChild(dialog);
                        const user = dialog.querySelector('[title="User Name"]');
                        const name = dialog.querySelector('[title="Variant Name"]');
                        dialog.querySelector('[title="Execute"]').addEventListener('click', () => {
                            window.__variant.user = user.value;
                            window.__variant.name = name.value;
                            dialog.innerHTML = `
                                <div id="variant-row" role="row">
                                    <div role="gridcell">STANDARD EU</div>
                                </div>`;
                            dialog.querySelector('#variant-row').addEventListener('dblclick', () => {
                                dialog.remove();
                            });
                        });
                    });
                </script>
                """
    )
    logs: list[tuple[str, str]] = []
    returned = hs_automation._apply_variant(
        page,
        page.main_frame,
        "STANDARD EU",
        lambda level, message: logs.append((level, message)),
    )
    result = page.evaluate("window.__variant")
    assert returned == page.main_frame
    assert result == {"user": "", "name": "STANDARD EU"}, (result, logs)
    assert not page.locator("#variant-dialog").count()


def test_apply_variant_accepts_accessible_get_variant_button(page) -> None:
        page.set_content(
                """
                <button aria-label="Get Variant...">Open</button>
                <script>
                    window.__clicked = false;
                    document.querySelector('button').addEventListener('click', () => {
                        window.__clicked = true;
                    });
                </script>
                """
        )
        hs_automation._click_by_title(page.main_frame, "Get Variant", timeout_ms=500)
        assert page.evaluate("window.__clicked") is True


def test_apply_variant_skips_when_no_variants_found(page) -> None:
    page.set_content(
        """
                <div title="Get Variant">Get Variant</div>
                <input title="Numbering Scheme" value="FHSCUKIM">
                <script>
                    document.querySelector('[title="Get Variant"]').addEventListener('click', () => {
                        const dialog = document.createElement('div');
                        dialog.id = 'variant-dialog';
                        dialog.setAttribute('role', 'dialog');
                        dialog.innerHTML = `
                            <input title="User Name" value="CURRENT">
                            <input title="Variant Name">
                            <button title="Execute">Execute</button>`;
                        document.body.appendChild(dialog);
                        dialog.querySelector('[title="Execute"]').addEventListener('click', () => {
                            const info = document.createElement('div');
                            info.id = 'info-dialog';
                            info.setAttribute('role', 'dialog');
                            info.textContent = 'InformationNo variants found for this selection x';
                            const close = document.createElement('button');
                            close.setAttribute('title', 'Close');
                            close.textContent = 'x';
                            close.addEventListener('click', () => {
                                info.remove();
                                dialog.remove();
                            });
                            info.appendChild(close);
                            document.body.appendChild(info);
                        });
                    });
                </script>
                """
    )
    logs: list[tuple[str, str]] = []
    try:
        hs_automation._apply_variant(
            page,
            page.main_frame,
            "STANDARD GB",
            lambda level, message: logs.append((level, message)),
        )
        raise AssertionError("missing variant should not apply")
    except hs_automation.HsMissingVariantError as exc:
        assert "STANDARD GB" in str(exc)
    assert not page.locator("#variant-dialog").count(), logs
    assert not page.locator("#info-dialog").count(), logs


def test_close_uses_cancel_without_extra_escape(page) -> None:
    page.set_content(dialog_html(include_start=False))
    hs_automation._close_dialog(page, page.main_frame)
    events = page.evaluate("window.__events")
    assert any(event.startswith("CLICK:Cancel") for event in events), events
    assert "KEY:Escape" not in events, events
    assert not page.locator("#webguiPopupWindow2").is_visible()


def test_concatenated_report_extracts_every_product() -> None:
    blob = (
        "TypMessage text"
        "Product 009424281 is classified with number 90271090 [ EDCHSCDEEX ]"
        "Product 011526161 is classified with number 90271090 [ EDCHSCDEEX ]"
        "Product 011613491 is classified with number 90271090 [ EDCHSCDEEX ]"
    )
    products = hs_automation._products_from_report(
        {"lines": [blob], "success": [blob], "products": []}
    )
    assert products == {"009424281", "011526161", "011613491"}, products


def test_virtualized_report_is_scrolled_before_continue(page) -> None:
    page.set_content(
        """
        <div id="C640_toolbar_btn8" role="button"
             title="Start Mass Classification (F8)">Start</div>
        <div id="webguiPopupWindow2" role="dialog">
          <div>Mass Classification</div>
          <label for="scheme">Numbering Scheme</label>
          <input id="scheme" title="Numbering Scheme" value="EDCHSCDEEX">
          <label for="tariff">Tariff Number</label>
          <input id="tariff" title="Tariff Number" value="">
        </div>
        <script>
          window.__events = [];
          window.__openReport = source => {
            window.__events.push('START:' + source);
            const mass = document.getElementById('webguiPopupWindow2');
            if (mass) mass.style.display = 'none';
            const report = document.createElement('div');
            report.id = 'SAPMSSY_REPORT';
            report.setAttribute('role', 'dialog');
            report.innerHTML = `
              <div id="userarealist1" style="height:24px; overflow:auto">
                <div class="lsAbapList__item">Product 355135851 is classified with number 90328900</div>
              </div>
              <div>0</div>
              <div>3</div>
              <div role="button" title="Continue (Enter)" id="continue">Continue</div>`;
            document.body.appendChild(report);
            const remaining = [
              'Product 355136431 is classified with number 90328900',
              'Product 355136441 is classified with number 90328900'
            ];
            const list = report.querySelector('#userarealist1');
            const reveal = () => {
              remaining.splice(0).forEach(text => {
                const row = document.createElement('div');
                row.className = 'lsAbapList__item';
                row.textContent = text;
                list.appendChild(row);
              });
            };
            list.addEventListener('scroll', reveal);
            document.addEventListener('keydown', ev => {
              if (ev.key === 'PageDown') reveal();
            });
            report.querySelector('#continue').addEventListener('click', () => report.remove());
          };
          document.addEventListener('click', ev => {
            const btn = ev.target.closest('[role=button]');
            if (!btn) return;
            window.__events.push('CLICK:' + (btn.getAttribute('title') || btn.id));
            if ((btn.getAttribute('title') || '').includes('Start Mass Classification')) {
              window.__openReport('CLICK');
            }
          });
        </script>
        """
    )
    logs: list[tuple[str, str]] = []
    hs_automation._fill_tariff_code(
        page, page.main_frame, "EDCHSCDEEX", "90328900", lambda *args: None
    )
    committed = hs_automation._commit_classification(
        page,
        page.main_frame,
        lambda level, msg: logs.append((level, msg)),
        ["355135851", "355136431", "355136441"],
    )
    assert committed == {"355135851", "355136431", "355136441"}, (committed, logs)
    assert any("confirmate=3/3" in msg for _, msg in logs), logs
    assert not page.locator("#SAPMSSY_REPORT").count()


def test_unexpected_report_product_aborts(page) -> None:
    html = dialog_html(include_start=True).replace(
        "Product 355136431", "Product 999999999"
    )
    page.set_content(html)
    frame = page.main_frame
    hs_automation._fill_tariff_code(
        page, frame, "EDCHSCDEEX", "90328900", lambda level, msg: None
    )
    try:
        hs_automation._commit_classification(
            page,
            frame,
            lambda level, msg: None,
            ["355135851", "355136431"],
        )
    except RuntimeError as exc:
        assert "nu apartin grupului" in str(exc), exc
    else:
        raise AssertionError("Unexpected SAP report product was accepted")


def test_group_reporting_is_complete() -> None:
    original = hs_automation._classify_group_via_multiple_selection
    events: list[dict] = []
    groups = [
        {
            "hs_code": "90328900",
            "products": [
                {"product": "355135851", "text": "A"},
                {"product": "355136431", "text": "B"},
            ],
        }
    ]

    def fake_group(page, frame, products, *args, **kwargs):
        return len(products), frame, {products[0]}

    hs_automation._classify_group_via_multiple_selection = fake_group
    try:
        summary = hs_automation._classify_groups(
            None,
            object(),
            groups,
            "EDCHSCDEEX",
            "STANDARD EU",
            True,
            threading.Event(),
            events.append,
            lambda level, msg: None,
        )
    finally:
        hs_automation._classify_group_via_multiple_selection = original

    group_end = [event for event in events if event.get("type") == "group-end"]
    materials = [event for event in events if event.get("type") == "material-end"]
    assert len(group_end) == 1, events
    assert group_end[0]["status"] == "partial" and not group_end[0]["ok"]
    assert [event["status"] for event in materials] == ["committed", "failed"]
    assert summary["materials_committed"] == 1
    assert summary["materials_failed"] == 1
    assert summary["materials_total_reported"] == 2

    safety_events: list[dict] = []

    def unsafe_group(page, frame, products, *args, **kwargs):
        raise hs_automation.HsSafetyError("unexpected product in SAP report")

    hs_automation._classify_group_via_multiple_selection = unsafe_group
    try:
        safety_summary = hs_automation._classify_groups(
            None,
            object(),
            [groups[0], {**groups[0], "hs_code": "90329000"}],
            "EDCHSCDEEX",
            "STANDARD EU",
            True,
            threading.Event(),
            safety_events.append,
            lambda level, msg: None,
        )
    finally:
        hs_automation._classify_group_via_multiple_selection = original

    assert safety_summary["safety_aborted"] is True
    assert safety_summary["skipped"] == 1
    assert len([e for e in safety_events if e.get("type") == "group-start"]) == 1


def main() -> int:
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        page = browser.new_page()

        run_case(page, include_start=True, unrelated_popup=True)
        print("PASS real Start click + explicit SAP report")

        run_case(page, include_start=False)
        print("PASS guarded single F8 fallback")

        test_hidden_dialog_does_not_block_selection(page)
        print("PASS hidden SAP dialogs do not block selection recovery")

        test_classifier_prefers_ready_frame_over_stale_frame(page)
        print("PASS ready SAP frame wins over stale frame")

        test_variant_dialog_can_be_found_in_child_frame(page)
        print("PASS Get Variant dialog can be found across frames")

        test_apply_variant_completes_selection_flow(page)
        print("PASS complete Get Variant selection flow")

        test_apply_variant_accepts_accessible_get_variant_button(page)
        print("PASS accessible Get Variant button locator")

        test_apply_variant_skips_when_no_variants_found(page)
        print("PASS missing Get Variant result is skippable")

        test_close_uses_cancel_without_extra_escape(page)
        print("PASS Cancel closes dialog without an extra Escape")

        test_virtualized_report_is_scrolled_before_continue(page)
        print("PASS virtualized SAP report is scrolled before Continue")

        test_unexpected_report_product_aborts(page)
        print("PASS unexpected SAP report products abort the group")

        browser.close()
    test_concatenated_report_extracts_every_product()
    print("PASS concatenated SAP report extracts every product")
    test_group_reporting_is_complete()
    print("PASS complete per-group and per-material reporting")
    print("RESULT= PASS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
