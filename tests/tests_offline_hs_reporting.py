"""Offline regression for persistent HS reports and factual material status."""

from __future__ import annotations

import json
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "server"))

from hs_reporting import HsReportStore


def main() -> int:
    with tempfile.TemporaryDirectory(prefix="hs-reporting-") as tmp:
        root = Path(tmp)
        store = HsReportStore(root)
        store.start(
            "abc123",
            kind="queue",
            system="FT1",
            items=[{"scheme": "EDCHSCDEEX", "variant": "STANDARD EU"}],
            allow_commit=True,
            auto_classify=True,
        )
        store.record(
            "abc123",
            {
                "type": "scheme-start",
                "index": 1,
                "scheme": "EDCHSCDEEX",
                "variant": "STANDARD EU",
            },
        )
        store.record(
            "abc123",
            {
                "type": "scheme-analysis",
                "index": 1,
                "scheme": "EDCHSCDEEX",
                "analysis": "analysis1",
                "source": "worklist.xlsx",
                "summary": {"ready": 1, "blocked": 1, "groups": 1},
                # Large analysis payloads must not leak into the timeline.
                "groups": [{"products": ["secret-large-payload"]}],
            },
        )
        store.record(
            "abc123",
            {
                "type": "material-end",
                "scheme": "EDCHSCDEEX",
                "product": "100001",
                "hs_code": "90328900",
                "status": "committed",
                "msg": "Mentinut in SAP",
            },
        )
        store.record(
            "abc123",
            {
                "type": "description-item-end",
                "material": "100001",
                "ok": True,
                "msg": "DE+EN confirmed",
            },
        )
        artifact = store.artifact_target("abc123", "remaining.xlsx")
        artifact.write_bytes(b"synthetic")
        remaining = {
            "ok": True,
            "file": artifact.name,
            "url": "/hs/reports/abc123/artifacts/remaining.xlsx",
            "original_products": 2,
            "removed_products": 1,
            "remaining_products": 1,
        }
        store.record(
            "abc123",
            {
                "type": "scheme-end",
                "index": 1,
                "scheme": "EDCHSCDEEX",
                "ok": True,
                "classify_summary": {"materials_committed": 1, "failed": 0},
                "description_summary": {"saved": 1, "total": 1, "failed": 0},
                "remaining_worklist": remaining,
            },
        )
        store.record(
            "abc123",
            {
                "type": "queue-end",
                "summary": {"completed": 1, "failed": 0, "skipped": 0},
            },
        )
        store.finish("abc123")

        report = store.get("abc123")
        assert report is not None
        assert report["status"] == "completed"
        assert report["schemes"][0]["remaining_worklist"] == remaining
        assert report["materials"] == [
            {
                "scheme": "EDCHSCDEEX",
                "product": "100001",
                "hs_code": "90328900",
                "hs_status": "committed",
                "hs_message": "Mentinut in SAP",
                "description_status": "confirmed",
                "description_message": "DE+EN confirmed",
            }
        ]
        serialized_events = str(report["events"])
        assert "secret-large-payload" not in serialized_events
        log_text = "\n".join(item["text"] for item in report["log"])
        assert "Start queue pe FT1" in log_text
        assert "Schema pornita. Varianta aplicata: STANDARD EU." in log_text
        assert "Schema finalizata." in log_text
        assert store.artifact_path("abc123", "remaining.xlsx") == artifact
        assert store.artifact_path("abc123", "not-allowed.xlsx") is None
        log_file = store.log_path("abc123")
        assert log_file is not None
        assert "Schema finalizata." in log_file.read_text(encoding="utf-8")

        # A fresh store instance must recover the report from disk.
        restored = HsReportStore(root).get("abc123")
        assert restored == report
        listed = HsReportStore(root).list()
        assert listed[0]["id"] == "abc123"
        assert set(listed[0]["schemes"][0]) == {"index", "scheme", "status"}

        warning_store = HsReportStore(root)
        warning_store.start(
            "warning1",
            kind="classify",
            system="FT6",
            items=[{"scheme": "EDCHSCINXX", "variant": "INDIA"}],
            allow_commit=True,
            auto_classify=True,
        )
        warning_store.record(
            "warning1",
            {
                "type": "job-end",
                "ok": True,
                "scheme": "EDCHSCINXX",
                "summary": {"materials_committed": 1},
                "remaining_worklist": {
                    "ok": True,
                    "warning": "one confirmed product was absent from the source",
                },
            },
        )
        warning_store.finish("warning1")
        warning = warning_store.get("warning1")
        assert warning is not None and warning["status"] == "warning"
        assert warning["traceability_errors"] == [
            "one confirmed product was absent from the source"
        ]

        warning_store.start(
            "empty1",
            kind="queue",
            system="FT6",
            items=[{"scheme": "EDCHSCZAXX", "variant": "STANDARD ZA"}],
        )
        warning_store.record(
            "empty1",
            {
                "type": "scheme-end",
                "index": 1,
                "scheme": "EDCHSCZAXX",
                "ok": True,
                "status": "no_data",
                "msg": "Varianta nu a returnat niciun rand (worklist gol).",
            },
        )
        warning_store.record(
            "empty1",
            {
                "type": "queue-end",
                "summary": {"completed": 0, "failed": 0, "no_data": 1},
            },
        )
        warning_store.finish("empty1")
        empty = warning_store.get("empty1")
        assert empty is not None
        assert empty["schemes"][0]["status"] == "no_data"
        assert empty["status"] == "completed"
        assert any("Worklist gol" in item["text"] for item in empty["log"])

        warning_store.start(
            "novar1",
            kind="queue",
            system="FT6",
            items=[{"scheme": "FHSCUKIM", "variant": "STANDARD GB"}],
        )
        warning_store.record(
            "novar1",
            {
                "type": "scheme-end",
                "index": 1,
                "scheme": "FHSCUKIM",
                "ok": True,
                "status": "no_variant",
                "msg": "Nicio varianta SAP gasita pentru 'STANDARD GB'.",
            },
        )
        warning_store.record(
            "novar1",
            {
                "type": "queue-end",
                "summary": {"completed": 0, "failed": 0, "no_variant": 1},
            },
        )
        warning_store.finish("novar1")
        novar = warning_store.get("novar1")
        assert novar is not None
        assert novar["schemes"][0]["status"] == "no_variant"
        assert novar["status"] == "completed"
        assert any("Varianta SAP lipsa" in item["text"] for item in novar["log"])

        warning_store.start(
            "interrupted1",
            kind="queue",
            system="FT6",
            items=[{"scheme": "EDCHSCUSIM", "variant": "STANDARD US"}],
        )
        warning_store.record(
            "interrupted1",
            {"type": "scheme-start", "index": 1, "scheme": "EDCHSCUSIM"},
        )
        recovered_store = HsReportStore(root)
        interrupted = recovered_store.get("interrupted1")
        assert interrupted is not None
        assert interrupted["status"] == "failed"
        assert interrupted["interrupted"] is True
        assert interrupted["schemes"][0]["status"] == "failed"
        assert interrupted["events"][-1]["type"] == "interrupted"

        recovered_store.record(
            "abc123",
            {
                "type": "description-item-end",
                "material": "100099",
                "ok": False,
                "status": "not_found",
                "msg": "NU a fost gasit in SAP",
            },
        )
        with_skip = recovered_store.get("abc123")
        assert with_skip is not None
        skipped = next(
            item for item in with_skip["materials"] if item["product"] == "100099"
        )
        assert skipped["description_status"] == "not_found"
        assert recovered_store.delete("warning1") is True
        assert recovered_store.get("warning1") is None
        assert not (root / "warning1").exists()
        remaining_count = recovered_store.clear()
        assert remaining_count >= 1
        assert recovered_store.list() == []
        assert list(root.iterdir()) == []

        legacy = root / "oldjob"
        legacy.mkdir()
        (legacy / "report.json").write_text(
            json.dumps({"id": "oldjob", "status": "completed", "events": []}),
            encoding="utf-8",
        )
        archived_store = HsReportStore(root)
        archived_store.start(
            "fresh1",
            kind="queue",
            system="FT6",
            items=[{"scheme": "EDCHSCDEEX", "variant": "STANDARD EU"}],
        )
        archived_store.finish("fresh1")
        assert not legacy.exists()
        assert (root / "arhiva" / "oldjob" / "report.json").is_file()
        assert archived_store.get("oldjob") is None
        archived_store.clear()
        assert archived_store.list() == []
        assert (root / "arhiva" / "oldjob" / "report.json").is_file()

    print("PASS persistent log reports, archive, and artifact allowlist")
    print("RESULT= PASS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
