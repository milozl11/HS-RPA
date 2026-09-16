"""Offline regression for the HS -> DE/EN -> next-scheme barrier."""

from __future__ import annotations

import sys
import threading
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "server"))

import hs_automation as automation

LANGUAGES = [
    {"code": "DE", "column": "J", "country_name": "Germany"},
    {"code": "EN", "column": "K", "country_name": "United Kingdom"},
]


class StubFrame:
    """Mirrors playwright Frame.page, used to pick the tab hosting the WebGUI."""

    def __init__(self, page):
        self.page = page


def description_plan() -> dict:
    return {
        "languages": ["DE", "EN"],
        "missing": [],
        "items": [
            {
                "material": "100001",
                "tecdoc": "10",
                "country_groups": ["EU", "GB"],
                "languages": ["DE", "EN"],
                "descriptions": {"DE": "Text DE 1", "EN": "Text EN 1"},
            },
            {
                "material": "100002",
                "tecdoc": "20",
                "country_groups": ["EU"],
                "languages": ["DE"],
                "descriptions": {"DE": "Text DE 2"},
            },
        ],
    }


def queue_payload() -> dict:
    return {
        "items": [
            {"scheme": "SCHEME1", "variant": "VARIANT1"},
            {"scheme": "SCHEME2", "variant": "VARIANT2"},
        ],
        "auto_classify": True,
        "allow_commit": True,
        "reference_path": "/tmp/reference.xlsx",
        "reference_sheet": None,
        "description_languages": LANGUAGES,
        "description_display_maintained": False,
        "description_url": "https://sap/#description",
        "classification_url": "https://sap/#classify",
    }


def case_queue_order_and_failure_gate() -> None:
    """The queue must run HS then descriptions, even if HS confirmation is partial."""
    names = [
        "HsReference",
        "_return_to_selection",
        "_apply_variant",
        "_set_numbering_scheme",
        "_set_display_all_products",
        "_execute_worklist",
        "_export_worklist",
        "read_worklist",
        "analyze",
        "build_description_plan",
        "_classify_groups",
        "_maintain_description_plan",
    ]
    originals = {name: getattr(automation, name) for name in names}
    sequence: list[str] = []
    display_values: list[bool] = []
    events: list[dict] = []

    def fake_analyze(worklist, reference, scheme):
        return {
            "scheme": scheme,
            "summary": {
                "worklist_rows": 1,
                "distinct_products": 1,
                "ready": 1,
                "blocked": 0,
                "groups": 1,
                "blocked_by_reason": {},
            },
            "groups": [
                {
                    "hs_code": "90328900",
                    "products": [{"product": f"P-{scheme}", "tecdoc": "10"}],
                }
            ],
            "blocked": [],
        }

    def fake_plan(groups, reference, languages):
        product = groups[0]["products"][0]
        return {
            "languages": ["DE", "EN"],
            "missing": [],
            "items": [
                {
                    "material": product["product"],
                    "tecdoc": product["tecdoc"],
                    "descriptions": {"DE": "DE", "EN": "EN"},
                }
            ],
            "summary": {"products": 1, "ready": 1, "missing": 0},
        }

    def fake_classify(
        page,
        frame,
        groups,
        scheme,
        variant,
        allow_commit,
        cancel,
        cb,
        log,
        blocked=None,
        display_all_products=False,
    ):
        sequence.append(f"hs:{scheme}")
        products = [
            str(row.get("product") or "").strip()
            for group in groups
            for row in group.get("products") or []
            if str(row.get("product") or "").strip()
        ]
        if scheme == "SCHEME1":
            return {
                "committed": 0,
                "failed": 1,
                "skipped": 0,
                "materials_committed": 0,
                "committed_products": [],
                "safety_aborted": False,
            }
        return {
            "committed": 1,
            "failed": 0,
            "skipped": 0,
            "materials_committed": 1,
            "committed_products": products,
            "safety_aborted": False,
        }

    def fake_descriptions(
        page,
        plan,
        languages,
        description_url,
        classification_url,
        credentials,
        cancel,
        cb,
        log,
        display_maintained=True,
        **kwargs,
    ):
        display_values.append(display_maintained)
        material = plan["items"][0]["material"]
        scheme = material.removeprefix("P-")
        sequence.append(f"desc:{scheme}")
        return (
            {
                "total": 1,
                "saved": 1,
                "failed": 0,
                "not_found": 0,
                "skipped": 0,
                "languages": ["DE", "EN"],
            },
            object(),
        )

    try:
        automation.HsReference = lambda path, sheet=None: object()
        automation._return_to_selection = lambda page, frame, log: frame
        automation._apply_variant = lambda page, frame, variant, log: frame
        automation._set_numbering_scheme = lambda page, frame, scheme, log: None
        automation._set_display_all_products = lambda page, frame, enabled, log: None
        automation._execute_worklist = lambda page, frame, log: None
        automation._export_worklist = lambda page, frame, log: Path(
            "/tmp/worklist.xlsx"
        )
        automation.read_worklist = lambda path: [{}]
        automation.analyze = fake_analyze
        automation.build_description_plan = fake_plan
        automation._classify_groups = fake_classify
        automation._maintain_description_plan = fake_descriptions

        session = automation.HsSession()
        session._do_queue(object(), object(), queue_payload(), events.append)
        assert sequence == ["hs:SCHEME1", "desc:SCHEME1", "hs:SCHEME2", "desc:SCHEME2"]
        assert display_values == [False, False]
        queue_end = next(event for event in events if event["type"] == "queue-end")
        assert queue_end["summary"]["completed"] == 2

        sequence.clear()
        events.clear()

        def fail_descriptions(*args, **kwargs):
            sequence.append("desc:SCHEME1")
            raise automation.HsSafetyError("DE/EN not confirmed")

        automation._maintain_description_plan = fail_descriptions
        session = automation.HsSession()
        session._do_queue(object(), object(), queue_payload(), events.append)
        assert sequence == ["hs:SCHEME1", "desc:SCHEME1"]
        queue_end = next(event for event in events if event["type"] == "queue-end")
        assert queue_end["summary"]["completed"] == 0
        assert queue_end["summary"]["failed"] == 1
        assert queue_end["summary"]["skipped"] == 1
    finally:
        for name, value in originals.items():
            setattr(automation, name, value)


def case_queue_skips_missing_variant() -> None:
    """A missing SAP variant must not abort the remaining schemes."""
    names = [
        "HsReference",
        "_return_to_selection",
        "_apply_variant",
        "_set_numbering_scheme",
        "_set_display_all_products",
        "_execute_worklist",
        "_export_worklist",
        "read_worklist",
        "analyze",
        "build_description_plan",
        "_classify_groups",
        "_maintain_description_plan",
    ]
    originals = {name: getattr(automation, name) for name in names}
    sequence: list[str] = []
    events: list[dict] = []

    def fake_analyze(worklist, reference, scheme):
        return {
            "scheme": scheme,
            "summary": {
                "worklist_rows": 1,
                "distinct_products": 1,
                "ready": 1,
                "blocked": 0,
                "groups": 1,
                "blocked_by_reason": {},
            },
            "groups": [
                {
                    "hs_code": "90328900",
                    "products": [{"product": f"P-{scheme}", "tecdoc": "10"}],
                }
            ],
            "blocked": [],
        }

    def fake_plan(groups, reference, languages):
        product = groups[0]["products"][0]
        return {
            "languages": ["DE", "EN"],
            "missing": [],
            "items": [
                {
                    "material": product["product"],
                    "tecdoc": product["tecdoc"],
                    "descriptions": {"DE": "DE", "EN": "EN"},
                }
            ],
            "summary": {"products": 1, "ready": 1, "missing": 0},
        }

    def fake_classify(*args, **kwargs):
        scheme = args[3]
        sequence.append(f"hs:{scheme}")
        return {
            "committed": 1,
            "failed": 0,
            "skipped": 0,
            "materials_committed": 1,
            "committed_products": [f"P-{scheme}"],
            "safety_aborted": False,
        }

    def fake_descriptions(*args, **kwargs):
        plan = args[1]
        scheme = plan["items"][0]["material"].removeprefix("P-")
        sequence.append(f"desc:{scheme}")
        return (
            {
                "total": 1,
                "saved": 1,
                "failed": 0,
                "not_found": 0,
                "skipped": 0,
                "languages": ["DE", "EN"],
            },
            object(),
        )

    def fake_variant(page, frame, variant, log):
        sequence.append(f"variant:{variant}")
        if variant == "VARIANT1":
            raise automation.HsMissingVariantError(
                "Nicio varianta SAP gasita pentru 'VARIANT1'."
            )
        return frame

    try:
        automation.HsReference = lambda path, sheet=None: object()
        automation._return_to_selection = lambda page, frame, log: frame
        automation._apply_variant = fake_variant
        automation._set_numbering_scheme = lambda page, frame, scheme, log: None
        automation._set_display_all_products = lambda page, frame, enabled, log: None
        automation._execute_worklist = lambda page, frame, log: None
        automation._export_worklist = lambda page, frame, log: Path(
            "/tmp/worklist.xlsx"
        )
        automation.read_worklist = lambda path: [{}]
        automation.analyze = fake_analyze
        automation.build_description_plan = fake_plan
        automation._classify_groups = fake_classify
        automation._maintain_description_plan = fake_descriptions

        session = automation.HsSession()
        session._do_queue(object(), object(), queue_payload(), events.append)
        assert sequence == [
            "variant:VARIANT1",
            "variant:VARIANT2",
            "hs:SCHEME2",
            "desc:SCHEME2",
        ]
        queue_end = next(event for event in events if event["type"] == "queue-end")
        assert queue_end["summary"]["completed"] == 1
        assert queue_end["summary"]["no_variant"] == 1
        assert queue_end["summary"]["failed"] == 0
        assert queue_end["summary"]["skipped"] == 0
        scheme_ends = [event for event in events if event["type"] == "scheme-end"]
        assert scheme_ends[0]["status"] == "no_variant"
        assert scheme_ends[0]["ok"] is True
    finally:
        for name, value in originals.items():
            setattr(automation, name, value)


def case_missing_variant_dialog_text_is_recognized() -> None:
    sap_text = (
        r"#SAPMSDYP10\5f 1,#SAPMSDYP10\5f 1-ph{}"
        "InformationNo variants found for this selection x"
    )
    assert automation._dialog_texts_indicate_missing_variant([sap_text])
    assert not automation._dialog_texts_indicate_missing_variant(
        ["Get Variant Variant Name Created by"]
    )


def case_do_descriptions_skips_hs() -> None:
    original_maintain = automation._maintain_description_plan
    events: list[dict] = []
    called: list[str] = []

    def fake_descriptions(*args, **kwargs):
        called.append("desc")
        return (
            {
                "total": 1,
                "saved": 1,
                "failed": 0,
                "not_found": 0,
                "skipped": 0,
                "languages": ["DE", "EN"],
            },
            object(),
        )

    try:
        automation._maintain_description_plan = fake_descriptions
        session = automation.HsSession()
        session._do_descriptions(
            object(),
            object(),
            {
                "scheme": "EDCHSCDEEX",
                "description_plan": description_plan(),
                "description_languages": LANGUAGES,
                "description_url": "https://sap/#description",
                "classification_url": "https://sap/#classify",
                "description_display_maintained": False,
            },
            events.append,
        )
        assert called == ["desc"]
        job_end = next(event for event in events if event["type"] == "job-end")
        assert job_end["ok"] is True
        assert job_end["summary"]["descriptions_only"] is True
    finally:
        automation._maintain_description_plan = original_maintain


def case_description_frame_accepts_standalone_selection_screen() -> None:
    """Description app shares Get Variant; Product + Multiple Selection is enough."""
    from playwright.sync_api import sync_playwright

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        page = browser.new_page()
        page.set_content(
            """
            <input title="Product">
            <div title="Get Variant">Get Variant</div>
            <div title="Execute">Execute</div>
            <div role="button" title="Multiple Selection" aria-label="Multiple Selection">MS</div>
            <div>Manage Customs Commercial Descriptions</div>
            """,
            wait_until="domcontentloaded",
        )
        frame = automation._commercial_description_frame(page, timeout_s=3)
        browser.close()
    assert frame is not None


def case_description_frame_accepts_screen_mentioning_numbering_scheme() -> None:
    """Live SAP shows a numbering-scheme label on the description screen too."""
    from playwright.sync_api import sync_playwright

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        page = browser.new_page()
        page.set_content(
            """
            <div>Manage Customs Commercial Descriptions</div>
            <div>Product-Specific Criteria</div>
            <input title="Logical System Group">
            <input title="Product">
            <input title="Identification for Product Master Worklist">
            <div>General Criteria</div>
            <div role="checkbox" aria-label="Display Maintained Products"></div>
            <div>Classification</div>
            <label>Type of Numbering Scheme:</label>
            <input title="Type of Numbering Scheme">
            <div role="button" title="Multiple Selection">MS</div>
            <div title="Execute">Execute</div>
            """,
            wait_until="domcontentloaded",
        )
        frame = automation._commercial_description_frame(page, timeout_s=3)
        browser.close()
    assert frame is not None


def case_description_frame_rejects_classifier_scheme_screen() -> None:
    from playwright.sync_api import sync_playwright

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        page = browser.new_page()
        page.set_content(
            """
            <input title="Product">
            <input title="Numbering Scheme">
            <input title="Identification for Product Master Worklist">
            <div title="Get Variant">Get Variant</div>
            <div title="Start Mass Classification">Start Mass Classification</div>
            """,
            wait_until="domcontentloaded",
        )
        try:
            automation._commercial_description_frame(page, timeout_s=1)
        except TimeoutError:
            browser.close()
            return
        browser.close()
        raise AssertionError("Classifier screen was accepted as description app")


def case_descriptions_found_in_second_browser_tab() -> None:
    """SAP can open the WebGUI transaction in another tab; drive that tab."""
    import tempfile

    from playwright.sync_api import sync_playwright

    shell = Path(tempfile.gettempdir()) / "hs_shell_only.html"
    webgui = Path(tempfile.gettempdir()) / "hs_webgui_tab.html"
    shell.write_text(
        "<!doctype html><html><body><div>Fiori shell</div></body></html>",
        encoding="utf-8",
    )
    webgui.write_text(
        """<!doctype html><html><body>
        <input title="Product">
        <div role="checkbox" aria-label="Display Maintained Products"></div>
        <div role="button" title="Multiple Selection">MS</div>
        </body></html>""",
        encoding="utf-8",
    )
    logs: list[tuple[str, str]] = []
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        context = browser.new_context()
        page = context.new_page()
        page.goto(shell.as_uri() + "#CustomsProduct-classify", wait_until="load")
        gui_tab = context.new_page()
        gui_tab.goto(webgui.as_uri(), wait_until="load")
        frame = automation._navigate_to_sap_app(
            page,
            shell.as_uri() + "#CustomsProduct-manageCustomsDescription",
            automation._commercial_description_frame,
            None,
            lambda level, msg: logs.append((level, msg)),
            "Manage Customs Commercial Descriptions",
            timeout_s=12,
            prefer_full_load=True,
        )
        found_on_gui_tab = frame.page is gui_tab
        browser.close()
    assert found_on_gui_tab, logs


def case_descriptions_use_full_deeplink_not_hash() -> None:
    """The descriptions app must be opened with a full load, like standalone."""
    import tempfile

    from playwright.sync_api import sync_playwright

    fixture = Path(tempfile.gettempdir()) / "hs_desc_fullload.html"
    fixture.write_text(
        """<!doctype html><html><body>
        <div id="app">shell</div>
        <script>
          window.__hashHops = 0;
          window.addEventListener('hashchange', () => { window.__hashHops++; });
          if (location.hash.indexOf('manageCustomsDescription') !== -1) {
            const box = document.createElement('input');
            box.setAttribute('title', 'Product');
            document.body.appendChild(box);
            const exec = document.createElement('div');
            exec.setAttribute('title', 'Execute');
            document.body.appendChild(exec);
            const worklist = document.createElement('input');
            worklist.setAttribute('title', 'Worklist');
            document.body.appendChild(worklist);
            const disp = document.createElement('div');
            disp.setAttribute('role', 'checkbox');
            disp.setAttribute('aria-label', 'Display Maintained Products');
            document.body.appendChild(disp);
          }
        </script>
        </body></html>""",
        encoding="utf-8",
    )
    logs: list[tuple[str, str]] = []
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        page = browser.new_page()
        page.goto(
            fixture.as_uri() + "#CustomsProduct-classify",
            wait_until="domcontentloaded",
        )
        target = (
            page.url.split("#", 1)[0]
            + "#CustomsProduct-manageCustomsDescription?sap-ui-tech-hint=GUI"
        )
        returned = automation._navigate_to_sap_app(
            page,
            target,
            automation._commercial_description_frame,
            None,
            lambda level, msg: logs.append((level, msg)),
            "Manage Customs Commercial Descriptions",
            timeout_s=12,
            prefer_full_load=True,
        )
        hops = page.evaluate("window.__hashHops")
        browser.close()
    assert returned is not None
    assert hops == 0, hops
    assert not any("hash Fiori" in msg for _, msg in logs), logs
    assert any("deep link" in msg.lower() for _, msg in logs), logs


def case_empty_webgui_falls_back_to_standalone_deeplink() -> None:
    """If Fiori hash leaves an empty WebGUI, reload like the standalone app."""
    import tempfile

    from playwright.sync_api import sync_playwright

    fixture = Path(tempfile.gettempdir()) / "hs_fiori_empty_reload.html"
    fixture.write_text(
        """<!doctype html><html><body>
        <div id="app">shell</div>
        <script>
          function addProduct() {
            if (document.querySelector('input[title="Product"]')) return;
            const box = document.createElement('input');
            box.setAttribute('title', 'Product');
            document.body.appendChild(box);
            const exec = document.createElement('div');
            exec.setAttribute('title', 'Execute');
            document.body.appendChild(exec);
            const worklist = document.createElement('input');
            worklist.setAttribute('title', 'Worklist');
            document.body.appendChild(worklist);
            const disp = document.createElement('div');
            disp.setAttribute('role', 'checkbox');
            disp.setAttribute('aria-label', 'Display Maintained Products');
            document.body.appendChild(disp);
            const ms = document.createElement('div');
            ms.setAttribute('role', 'button');
            ms.setAttribute('title', 'Multiple Selection');
            document.body.appendChild(ms);
          }
          window.addEventListener('load', () => {
            if (location.hash.indexOf('manageCustomsDescription') !== -1) addProduct();
          });
        </script>
        </body></html>""",
        encoding="utf-8",
    )
    logs: list[tuple[str, str]] = []
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        page = browser.new_page()
        page.goto(
            fixture.as_uri() + "#CustomsProduct-classify",
            wait_until="domcontentloaded",
        )
        target = (
            page.url.split("#", 1)[0]
            + "#CustomsProduct-manageCustomsDescription?sap-ui-tech-hint=GUI"
        )
        returned = automation._navigate_to_sap_app(
            page,
            target,
            automation._commercial_description_frame,
            None,
            lambda level, msg: logs.append((level, msg)),
            "Manage Customs Commercial Descriptions",
            timeout_s=16,
        )
        browser.close()
    assert returned is not None
    assert any("hash Fiori" in msg for _, msg in logs), logs
    assert any(
        "standalone" in msg.lower() or "reload" in msg.lower() for _, msg in logs
    ), logs


def case_fiori_hash_navigation_skips_full_reload() -> None:
    """Same-document Fiori hops must not wait on a full page.goto."""
    import tempfile

    from playwright.sync_api import sync_playwright

    html = """<!doctype html><html><body>
            <div id="app">Classify Products</div>
            <script>
              window.__hashes = [];
              window.addEventListener('hashchange', () => {
                window.__hashes.push(location.hash);
                document.getElementById('app').textContent =
                  'Manage Customs Commercial Descriptions';
                const box = document.createElement('input');
                box.setAttribute('title', 'Product');
                document.body.appendChild(box);
                const exec = document.createElement('div');
                exec.setAttribute('title', 'Execute');
                document.body.appendChild(exec);
              });            </script>
            </body></html>"""
    fixture = Path(tempfile.gettempdir()) / "hs_fiori_nav.html"
    fixture.write_text(html, encoding="utf-8")
    logs: list[tuple[str, str]] = []
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        page = browser.new_page()
        page.goto(
            fixture.as_uri() + "#CustomsProduct-classify", wait_until="domcontentloaded"
        )
        returned = automation._navigate_to_sap_app(
            page,
            page.url.split("#", 1)[0]
            + "#CustomsProduct-manageCustomsDescription?sap-ui-tech-hint=GUI",
            automation._commercial_description_frame,
            None,
            lambda level, msg: logs.append((level, msg)),
            "Manage Customs Commercial Descriptions",
            timeout_s=8,
        )
        hashes = page.evaluate("window.__hashes")
        browser.close()
    assert returned is not None
    assert any("manageCustomsDescription" in value for value in hashes), hashes
    assert any("hash Fiori" in msg for _, msg in logs), logs
    assert not any("goto" in msg.lower() and "esuat" in msg.lower() for _, msg in logs)


def case_classify_starts_descriptions_after_partial_hs() -> None:
    """Manual classify must start DE/EN even when SAP report harvest is incomplete."""
    sequence: list[str] = []
    events: list[dict] = []
    original_classify = automation._classify_groups
    original_maintain = automation._maintain_description_plan

    def fake_classify(*args, **kwargs):
        sequence.append("hs")
        return {
            "committed": 0,
            "failed": 1,
            "skipped": 0,
            "materials_committed": 21,
            "committed_products": ["100001"],
            "safety_aborted": False,
        }

    def fake_descriptions(page, plan, *args, **kwargs):
        sequence.append("desc")
        assert [item["material"] for item in plan["items"]] == ["100001", "100002"]
        return ({"saved": 2, "failed": 0}, object())

    try:
        automation._classify_groups = fake_classify
        automation._maintain_description_plan = fake_descriptions
        session = automation.HsSession()
        session._do_classify(
            object(),
            object(),
            {
                "groups": [
                    {"products": [{"product": "100001"}, {"product": "100002"}]}
                ],
                "scheme": "EDCHSCDEEX",
                "variant": "STANDARD EU",
                "blocked": [],
                "allow_commit": True,
                "description_plan": description_plan(),
                "description_languages": LANGUAGES,
                "description_url": "https://sap/#description",
                "classification_url": "https://sap/#classify",
                "description_display_maintained": False,
            },
            events.append,
        )
    finally:
        automation._classify_groups = original_classify
        automation._maintain_description_plan = original_maintain

    assert sequence == ["hs", "desc"], sequence
    job_end = next(event for event in events if event["type"] == "job-end")
    assert job_end["ok"] is True
    logs = [event["msg"] for event in events if event.get("type") == "log"]
    assert any("Continui cu descrierile comerciale" in msg for msg in logs), logs


def case_country_grid_id_matches_sap_cell_ids() -> None:
    """SAP marks role=grid on '<id>-content' while cells keep the base id."""
    import sap_automation
    from playwright.sync_api import sync_playwright

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        page = browser.new_page()
        page.set_content(
            """
            <div id="M1:46:3">
              <div id="M1:46:3-content" role="grid">
                <div role="columnheader">Lang.</div>
                <div role="columnheader">Descriptn</div>
                <div role="row">
                  <div role="gridcell" id="M1:46:3[1,1]"></div>
                  <div role="gridcell" id="M1:46:3[1,2]"></div>
                </div>
              </div>
            </div>
            <div id="C131-content" role="grid" aria-label="Ctries">
              <div role="row"><div role="gridcell">DE</div></div>
            </div>
            """,
            wait_until="domcontentloaded",
        )
        ids = sap_automation._country_grid_ids(page.main_frame)
        browser.close()
    assert ids["countryGrid"] == "M1:46:3", ids
    assert ids["countryListGrid"] == "C131-content", ids


def case_failed_material_recovers_selection_screen() -> None:
    """A failed material must not cascade into the whole worklist."""
    original_navigate = automation._navigate_to_sap_app
    original_frame = automation._commercial_description_frame
    original_process = automation._process_description_material
    original_display = automation._set_description_display_maintained
    original_recover = automation._recover_description_screen
    processed: list[str] = []
    recovered: list[str] = []

    def fake_process(page, frame, material, entries, log):
        processed.append(material)
        if len(processed) == 1:
            raise RuntimeError("Locator.click: Timeout 5000ms exceeded.")
        return "ok"

    try:
        automation._navigate_to_sap_app = (
            lambda page, url, finder, creds, log, label, **kw: StubFrame(page)
        )
        automation._commercial_description_frame = lambda page, timeout_s=10: StubFrame(
            page
        )
        automation._process_description_material = fake_process
        automation._set_description_display_maintained = lambda frame, want, log: None
        automation._recover_description_screen = lambda page, log: recovered.append(
            "recover"
        )
        summary, _ = automation._maintain_description_plan(
            object(),
            description_plan(),
            LANGUAGES,
            "https://sap/#description",
            "https://sap/#classify",
            None,
            threading.Event(),
            lambda ev: None,
            lambda level, msg: None,
        )
    finally:
        automation._navigate_to_sap_app = original_navigate
        automation._commercial_description_frame = original_frame
        automation._process_description_material = original_process
        automation._set_description_display_maintained = original_display
        automation._recover_description_screen = original_recover

    assert processed == ["100001", "100002"], processed
    assert recovered == ["recover"], recovered
    assert summary["saved"] == 1 and summary["failed"] == 1, summary


def bulk_plan() -> dict:
    """Two TECDOCs: one EU+GB pair, one EU-only, plus a second EU-only TECDOC."""
    return {
        "languages": ["DE", "EN"],
        "missing": [],
        "items": [
            {
                "material": "100001",
                "tecdoc": "10",
                "languages": ["DE", "EN"],
                "descriptions": {"DE": "Text DE 10", "EN": "Text EN 10"},
            },
            {
                "material": "100002",
                "tecdoc": "10",
                "languages": ["DE"],
                "descriptions": {"DE": "Text DE 10"},
            },
            {
                "material": "100003",
                "tecdoc": "10",
                "languages": ["DE", "EN"],
                "descriptions": {"DE": "Text DE 10", "EN": "Text EN 10"},
            },
            {
                "material": "100004",
                "tecdoc": "20",
                "languages": ["DE"],
                "descriptions": {"DE": "Text DE 20"},
            },
        ],
    }


def case_items_group_by_tecdoc_and_languages() -> None:
    """EU+GB and EU-only products of one TECDOC must be separate SAP passes."""
    language_by_code = {lang["code"]: lang for lang in LANGUAGES}
    groups, invalid = automation._group_description_items(
        bulk_plan()["items"], language_by_code, ["DE", "EN"]
    )
    assert invalid == [], invalid
    keys = [
        (group["tecdoc"], [entry[0] for entry in group["lang_entries"]])
        for group in groups
    ]
    assert keys == [("10", ["DE", "EN"]), ("10", ["DE"]), ("20", ["DE"])], keys
    members = [[item["material"] for _, item in group["members"]] for group in groups]
    assert members == [["100001", "100003"], ["100002"], ["100004"]], members


def case_bulk_maintenance_uses_one_pass_per_group() -> None:
    """Multi-product groups go through the mass flow, singles stay per material."""
    original_navigate = automation._navigate_to_sap_app
    original_frame = automation._commercial_description_frame
    original_bulk = automation._maintain_description_bulk
    original_process = automation._process_description_material
    original_display = automation._set_description_display_maintained
    bulk_calls: list[tuple[tuple[str, ...], tuple[str, ...]]] = []
    single_calls: list[str] = []
    events: list[dict] = []

    def fake_bulk(page, materials, lang_entries, label, log):
        bulk_calls.append((tuple(materials), tuple(entry[0] for entry in lang_entries)))

    def fake_process(page, frame, material, entries, log):
        single_calls.append(material)
        return "ok"

    try:
        automation._navigate_to_sap_app = (
            lambda page, url, finder, creds, log, label, **kw: StubFrame(page)
        )
        automation._commercial_description_frame = lambda page, timeout_s=10: StubFrame(
            page
        )
        automation._maintain_description_bulk = fake_bulk
        automation._process_description_material = fake_process
        automation._set_description_display_maintained = lambda frame, want, log: None
        summary, _ = automation._maintain_description_plan(
            object(),
            bulk_plan(),
            LANGUAGES,
            "https://sap/#description",
            "https://sap/#classify",
            None,
            threading.Event(),
            events.append,
            lambda level, msg: None,
        )
    finally:
        automation._navigate_to_sap_app = original_navigate
        automation._commercial_description_frame = original_frame
        automation._maintain_description_bulk = original_bulk
        automation._process_description_material = original_process
        automation._set_description_display_maintained = original_display

    assert bulk_calls == [(("100001", "100003"), ("DE", "EN"))], bulk_calls
    assert single_calls == ["100002", "100004"], single_calls
    assert summary["saved"] == 4, summary
    saved = [
        ev["material"]
        for ev in events
        if ev.get("type") == "description-item-end" and ev.get("ok")
    ]
    assert sorted(saved) == ["100001", "100002", "100003", "100004"], saved


def case_bulk_failure_falls_back_to_single_materials() -> None:
    """A failed mass pass must retry its products one by one."""
    original_navigate = automation._navigate_to_sap_app
    original_frame = automation._commercial_description_frame
    original_bulk = automation._maintain_description_bulk
    original_process = automation._process_description_material
    original_display = automation._set_description_display_maintained
    original_recover = automation._recover_description_screen
    single_calls: list[str] = []

    def failing_bulk(page, materials, lang_entries, label, log):
        raise RuntimeError("SAP nu a confirmat salvarea in masa")

    def fake_process(page, frame, material, entries, log):
        single_calls.append(material)
        return "ok"

    try:
        automation._navigate_to_sap_app = (
            lambda page, url, finder, creds, log, label, **kw: StubFrame(page)
        )
        automation._commercial_description_frame = lambda page, timeout_s=10: StubFrame(
            page
        )
        automation._maintain_description_bulk = failing_bulk
        automation._process_description_material = fake_process
        automation._set_description_display_maintained = lambda frame, want, log: None
        automation._recover_description_screen = lambda page, log: None
        summary, _ = automation._maintain_description_plan(
            object(),
            bulk_plan(),
            LANGUAGES,
            "https://sap/#description",
            "https://sap/#classify",
            None,
            threading.Event(),
            lambda ev: None,
            lambda level, msg: None,
        )
    finally:
        automation._navigate_to_sap_app = original_navigate
        automation._commercial_description_frame = original_frame
        automation._maintain_description_bulk = original_bulk
        automation._process_description_material = original_process
        automation._set_description_display_maintained = original_display
        automation._recover_description_screen = original_recover

    assert single_calls == ["100001", "100003", "100002", "100004"], single_calls
    assert summary["saved"] == 4, summary


def main() -> int:
    assert (
        automation.commercial_description_url(
            "https://cift1.example/sap/bc/ui2/flp?client=100#CustomsProduct-classify?x=1"
        )
        == "https://cift1.example/sap/bc/ui2/flp?client=100"
        "#CustomsProduct-manageCustomsDescription?sap-ui-tech-hint=GUI"
    )

    groups = [
        {"products": [{"product": "1"}, {"product": "2"}]},
        {"products": [{"product": "3"}]},
    ]
    complete = {
        "committed": 2,
        "failed": 0,
        "skipped": 0,
        "materials_committed": 3,
        "safety_aborted": False,
    }
    assert automation._classification_completion_error(complete, groups, True) == ""
    partial = {**complete, "materials_committed": 2}
    assert (
        "nu a confirmat toate materialele"
        in automation._classification_completion_error(partial, groups, True)
    )
    aborted = {**complete, "safety_aborted": True, "abort_reason": "selectie invalida"}
    assert "selectie invalida" in automation._classification_completion_error(
        aborted, groups, True
    )
    assert "dry-run" in automation._classification_completion_error(
        complete, groups, False
    )
    case_queue_order_and_failure_gate()
    case_queue_skips_missing_variant()
    case_missing_variant_dialog_text_is_recognized()
    case_do_descriptions_skips_hs()
    case_classify_starts_descriptions_after_partial_hs()
    case_description_frame_accepts_standalone_selection_screen()
    case_description_frame_accepts_screen_mentioning_numbering_scheme()
    case_description_frame_rejects_classifier_scheme_screen()
    case_empty_webgui_falls_back_to_standalone_deeplink()
    case_descriptions_use_full_deeplink_not_hash()
    case_descriptions_found_in_second_browser_tab()
    case_fiori_hash_navigation_skips_full_reload()
    case_country_grid_id_matches_sap_cell_ids()
    case_failed_material_recovers_selection_screen()
    case_items_group_by_tecdoc_and_languages()
    case_bulk_maintenance_uses_one_pass_per_group()
    case_bulk_failure_falls_back_to_single_materials()

    original_navigate = automation._navigate_to_sap_app
    original_frame = automation._commercial_description_frame
    original_process = automation._process_description_material
    original_display = automation._set_description_display_maintained
    calls: list[tuple] = []
    display_values: list[bool] = []
    events: list[dict] = []

    def fake_navigate(page, url, finder, credentials, log, label, timeout_s=180, **kw):
        calls.append(("navigate", label, url))
        return StubFrame(page)

    def fake_process(page, frame, material, entries, log):
        calls.append(("process", material, entries))
        return "ok"

    try:
        automation._navigate_to_sap_app = fake_navigate
        automation._commercial_description_frame = lambda page, timeout_s=10: StubFrame(
            page
        )
        automation._process_description_material = fake_process
        automation._set_description_display_maintained = lambda frame, want, log: (
            display_values.append(want)
        )
        summary, _ = automation._maintain_description_plan(
            object(),
            description_plan(),
            LANGUAGES,
            "https://sap/#description",
            "https://sap/#classify",
            None,
            threading.Event(),
            events.append,
            lambda level, msg: None,
            display_maintained=False,
        )
        assert display_values == [False]
        assert summary["saved"] == 2 and summary["failed"] == 0
        process_calls = [call for call in calls if call[0] == "process"]
        assert [call[1] for call in process_calls] == ["100001", "100002"]
        assert [entry[0] for entry in process_calls[0][2]] == ["DE", "EN"]
        assert [entry[0] for entry in process_calls[1][2]] == ["DE"]
        assert calls[-1][0:2] == ("navigate", "Classify Products")
        assert events[-1]["type"] == "description-end" and events[-1]["ok"]

        # A missing SAP material is recorded and skipped; remaining work continues.
        calls.clear()
        events.clear()

        def fail_second(page, frame, material, entries, log):
            calls.append(("process", material, entries))
            return "not_found" if material == "100002" else "ok"

        automation._process_description_material = fail_second
        skipped_summary, _ = automation._maintain_description_plan(
            object(),
            description_plan(),
            LANGUAGES,
            "https://sap/#description",
            "https://sap/#classify",
            None,
            threading.Event(),
            events.append,
            lambda level, msg: None,
        )
        assert skipped_summary["saved"] == 1
        assert skipped_summary["not_found"] == 1
        assert skipped_summary["failed"] == 0
        assert [call[1] for call in calls if call[0] == "process"] == [
            "100001",
            "100002",
        ]
        assert any(call[0:2] == ("navigate", "Classify Products") for call in calls)
        assert events[-1]["type"] == "description-end"
        assert events[-1]["ok"] is True
        assert events[-1].get("warning") is True
        item_ends = [
            event for event in events if event["type"] == "description-item-end"
        ]
        assert item_ends[-1]["status"] == "not_found"
        assert item_ends[-1]["ok"] is False

        calls.clear()
        events.clear()
        empty_summary, _ = automation._maintain_description_plan(
            object(),
            {
                "languages": [],
                "missing": [],
                "skipped": [{"material": "1"}],
                "items": [],
            },
            LANGUAGES,
            "https://sap/#description",
            "https://sap/#classify",
            None,
            threading.Event(),
            events.append,
            lambda level, msg: None,
        )
        assert empty_summary["saved"] == 0
        assert not any(call[0] == "process" for call in calls)
        assert not any(
            call[1] == "Manage Customs Commercial Descriptions" for call in calls
        )
        assert events[-1]["type"] == "description-end" and events[-1]["ok"]
    finally:
        automation._navigate_to_sap_app = original_navigate
        automation._commercial_description_frame = original_frame
        automation._process_description_material = original_process
        automation._set_description_display_maintained = original_display

    print("PASS HS descriptions continue after missing SAP material")
    print("RESULT= PASS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
