"""Persistent, log-oriented reporting for the integrated HS workflow."""

from __future__ import annotations

import json
import re
import shutil
import threading
import time
from datetime import datetime
from pathlib import Path
from typing import Any

ARCHIVE_NAME = "arhiva"
REPORT_FORMAT = "log-v1"


def _now() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def _clean_id(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9_-]+", "", str(value or ""))
    if not cleaned:
        raise ValueError("Report id invalid.")
    return cleaned


def _first_line(value: Any, limit: int = 500) -> str:
    lines = str(value or "").splitlines()
    return lines[0][:limit] if lines else ""


def _int(value: Any) -> int:
    try:
        return int(value or 0)
    except (TypeError, ValueError):
        return 0


def _stamp(value: str = "") -> str:
    raw = value or _now()
    return raw.replace("T", " ")[:19]


def _counts(payload: Any, *keys: str) -> str:
    data = payload if isinstance(payload, dict) else {}
    parts = []
    for key in keys:
        if key in data and data.get(key) is not None:
            parts.append(f"{key}={data.get(key)}")
    return ", ".join(parts)


class HsReportStore:
    """Keep a chronological operator log plus compact scheme/material facts."""

    _PERSIST_EVENTS = {
        "queue-start",
        "job-start",
        "scheme-start",
        "scheme-analysis",
        "analysis",
        "remaining-worklist",
        "group-end",
        "description-preflight",
        "description-end",
        "scheme-end",
        "job-end",
        "queue-end",
        "fatal",
    }

    def __init__(self, root: str | Path):
        self.root = Path(root)
        self.root.mkdir(parents=True, exist_ok=True)
        self._lock = threading.RLock()
        self._active: dict[str, dict] = {}
        self._archive_legacy_reports()
        self._recover_interrupted_reports()

    def start(
        self,
        report_id: str,
        *,
        kind: str,
        system: str,
        items: list[dict] | None = None,
        allow_commit: bool = False,
        auto_classify: bool = False,
        source_file: str = "",
    ) -> dict:
        report_id = _clean_id(report_id)
        schemes = []
        for index, item in enumerate(items or [], start=1):
            schemes.append(
                {
                    "index": index,
                    "scheme": str(item.get("scheme") or ""),
                    "variant": str(item.get("variant") or ""),
                    "status": "pending",
                    "source_file": source_file if len(items or []) == 1 else "",
                    "analysis": None,
                    "classification": None,
                    "descriptions": None,
                    "remaining_worklist": None,
                    "error": "",
                }
            )
        created = _now()
        report = {
            "id": report_id,
            "format": REPORT_FORMAT,
            "kind": kind,
            "system": system,
            "mode": "commit" if allow_commit else "dry_run",
            "auto_classify": bool(auto_classify),
            "status": "running",
            "interrupted": False,
            "created_at": created,
            "finished_at": "",
            "summary": None,
            "error": "",
            "traceability_errors": [],
            "schemes": schemes,
            "materials": [],
            "events": [],
            "log": [
                {
                    "sequence": 1,
                    "timestamp": created,
                    "level": "info",
                    "scheme": "",
                    "text": self._start_line(
                        kind=kind,
                        system=system,
                        allow_commit=allow_commit,
                        auto_classify=auto_classify,
                        schemes=schemes,
                        source_file=source_file,
                    ),
                }
            ],
        }
        with self._lock:
            self._active[report_id] = report
            self._write(report)
        return report

    def record(self, report_id: str, event: dict) -> dict | None:
        report_id = _clean_id(report_id)
        with self._lock:
            report = self._load(report_id)
            if report is None:
                return None
            event_record = self._event_record(report, event)
            report["events"].append(event_record)
            log_entry = self._log_entry(report, event, event_record)
            if log_entry:
                report.setdefault("log", []).append(log_entry)
            self._apply_event(report, event)
            if (
                event.get("type") in self._PERSIST_EVENTS
                or len(report["events"]) % 25 == 0
            ):
                self._write(report)
            return report

    def finish(self, report_id: str) -> dict | None:
        report_id = _clean_id(report_id)
        with self._lock:
            report = self._load(report_id)
            if report is None:
                return None
            if report["status"] == "running":
                report["status"] = (
                    "warning" if report.get("traceability_errors") else "completed"
                )
            if not report.get("finished_at"):
                report["finished_at"] = _now()
            self._append_unique_log(
                report,
                level="info" if report["status"] != "failed" else "error",
                text=f"Raport inchis cu status {report['status']}.",
            )
            self._write(report)
            return report

    def get(self, report_id: str) -> dict | None:
        report_id = _clean_id(report_id)
        with self._lock:
            report = self._load(report_id)
            if report is None:
                return None
            # Return a detached value so a Flask serializer cannot observe a
            # concurrent event mutation half-way through a response.
            return json.loads(json.dumps(report, ensure_ascii=False))

    def list(self, limit: int = 100) -> list[dict]:
        reports = []
        with self._lock:
            paths = sorted(
                (
                    path
                    for path in self.root.glob("*/report.json")
                    if path.parent.name.lower() != ARCHIVE_NAME
                ),
                key=lambda path: path.stat().st_mtime_ns,
                reverse=True,
            )
            for path in paths[: max(1, min(int(limit), 500))]:
                try:
                    report = json.loads(path.read_text(encoding="utf-8"))
                except (OSError, json.JSONDecodeError):
                    continue
                log_entries = report.get("log") or []
                headline = ""
                if log_entries:
                    headline = str(log_entries[-1].get("text") or "")
                reports.append(
                    {
                        "id": report.get("id"),
                        "kind": report.get("kind"),
                        "system": report.get("system"),
                        "mode": report.get("mode"),
                        "status": report.get("status"),
                        "created_at": report.get("created_at"),
                        "finished_at": report.get("finished_at"),
                        "error": report.get("error"),
                        "traceability_errors": report.get("traceability_errors"),
                        "headline": headline,
                        "schemes": [
                            {
                                "index": scheme.get("index"),
                                "scheme": scheme.get("scheme"),
                                "status": scheme.get("status"),
                            }
                            for scheme in report.get("schemes") or []
                        ],
                    }
                )
        return reports

    def stats(self) -> list[dict]:
        """Aggregate HS and description maintenance per calendar month."""
        buckets: dict[str, dict] = {}
        with self._lock:
            paths = (
                path
                for path in self.root.glob("*/report.json")
                if path.parent.name.lower() != ARCHIVE_NAME
            )
            for path in paths:
                try:
                    report = json.loads(path.read_text(encoding="utf-8"))
                except (OSError, json.JSONDecodeError):
                    continue
                created = str(report.get("created_at") or "")[:7]
                if len(created) != 7:
                    continue
                bucket = buckets.setdefault(
                    created,
                    {
                        "month": created,
                        "runs": 0,
                        "hs_committed": 0,
                        "hs_failed": 0,
                        "desc_saved": 0,
                        "desc_failed": 0,
                    },
                )
                bucket["runs"] += 1
                for scheme in report.get("schemes") or []:
                    classification = scheme.get("classification") or {}
                    descriptions = scheme.get("descriptions") or {}
                    bucket["hs_committed"] += _int(
                        classification.get("materials_committed")
                    )
                    bucket["hs_failed"] += _int(classification.get("materials_failed"))
                    bucket["desc_saved"] += _int(descriptions.get("saved"))
                    bucket["desc_failed"] += _int(descriptions.get("failed")) + _int(
                        descriptions.get("not_found")
                    )
        return [buckets[key] for key in sorted(buckets)]

    def artifact_target(self, report_id: str, filename: str) -> Path:
        report_id = _clean_id(report_id)
        safe_name = re.sub(r"[^A-Za-z0-9._-]+", "_", Path(filename).name)
        if not safe_name:
            raise ValueError("Artifact filename invalid.")
        folder = self.root / report_id
        folder.mkdir(parents=True, exist_ok=True)
        return folder / safe_name

    def delete(self, report_id: str) -> bool:
        """Remove one persisted report folder and drop it from memory."""
        report_id = _clean_id(report_id)
        if report_id.lower() == ARCHIVE_NAME:
            return False
        with self._lock:
            folder = self.root / report_id
            existed = report_id in self._active or folder.exists()
            self._active.pop(report_id, None)
            if folder.exists():
                shutil.rmtree(folder, ignore_errors=True)
            return existed and not folder.exists()

    def clear(self) -> int:
        """Delete active reports. The archive folder is left untouched."""
        removed = 0
        with self._lock:
            self._active.clear()
            for folder in list(self.root.iterdir()):
                if not folder.is_dir() or folder.name.lower() == ARCHIVE_NAME:
                    continue
                shutil.rmtree(folder, ignore_errors=True)
                if not folder.exists():
                    removed += 1
        return removed

    def artifact_path(self, report_id: str, filename: str) -> Path | None:
        report = self.get(report_id)
        if report is None:
            return None
        wanted = Path(filename).name
        allowed = {
            Path(str(scheme.get("remaining_worklist", {}).get("file") or "")).name
            for scheme in report.get("schemes") or []
            if scheme.get("remaining_worklist")
        }
        if wanted not in allowed:
            return None
        path = self.root / _clean_id(report_id) / wanted
        return path if path.is_file() else None

    def log_path(self, report_id: str) -> Path | None:
        path = self.root / _clean_id(report_id) / "report.log"
        return path if path.is_file() else None

    def _path(self, report_id: str) -> Path:
        return self.root / report_id / "report.json"

    def _load(self, report_id: str) -> dict | None:
        if report_id in self._active:
            return self._active[report_id]
        path = self._path(report_id)
        if not path.is_file():
            return None
        try:
            report = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return None
        report.setdefault("log", [])
        report.setdefault("format", REPORT_FORMAT)
        self._active[report_id] = report
        return report

    def _write(self, report: dict) -> None:
        path = self._path(report["id"])
        path.parent.mkdir(parents=True, exist_ok=True)
        temp = path.with_suffix(".tmp")
        temp.write_text(
            json.dumps(report, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        # Cloud-synced folders (OneDrive) or AV scanners can hold a transient
        # lock on the destination during rename; retry briefly instead of
        # surfacing noise on every event for a lock that clears in ms.
        last_exc: OSError | None = None
        for attempt in range(6):
            try:
                temp.replace(path)
                last_exc = None
                break
            except OSError as exc:
                last_exc = exc
                time.sleep(0.05 * (2**attempt))
        if last_exc is not None:
            raise last_exc
        self._write_log_file(path.parent / "report.log", report.get("log") or [])

    def _write_log_file(self, path: Path, entries: list[dict]) -> None:
        lines = []
        for entry in entries:
            scheme = str(entry.get("scheme") or "").strip()
            prefix = f"[{scheme}] " if scheme else ""
            lines.append(
                f"{_stamp(str(entry.get('timestamp') or ''))}  "
                f"{str(entry.get('level') or 'info').upper():<5}  "
                f"{prefix}{entry.get('text') or ''}".rstrip()
            )
        payload = "\n".join(lines)
        if payload:
            payload += "\n"
        temp = path.with_suffix(".tmp")
        temp.write_text(payload, encoding="utf-8")
        last_exc: OSError | None = None
        for attempt in range(6):
            try:
                temp.replace(path)
                return
            except OSError as exc:
                last_exc = exc
                time.sleep(0.05 * (2**attempt))
        if last_exc is not None:
            raise last_exc

    def _archive_legacy_reports(self) -> None:
        """Move pre-LOG job folders into reports-hs/arhiva and keep them."""
        archive_root = self.root / ARCHIVE_NAME
        for folder in list(self.root.iterdir()):
            if not folder.is_dir() or folder.name.lower() == ARCHIVE_NAME:
                continue
            report_path = folder / "report.json"
            if report_path.is_file():
                try:
                    data = json.loads(report_path.read_text(encoding="utf-8"))
                except (OSError, json.JSONDecodeError):
                    data = {}
                if data.get("format") == REPORT_FORMAT:
                    continue
            archive_root.mkdir(parents=True, exist_ok=True)
            dest = archive_root / folder.name
            if dest.exists():
                dest = archive_root / (
                    f"{folder.name}-{datetime.now().strftime('%Y%m%d%H%M%S')}"
                )
            try:
                shutil.move(str(folder), str(dest))
            except OSError:
                continue

    def _recover_interrupted_reports(self) -> None:
        """Close jobs left running by a previous application process."""
        for path in self.root.glob("*/report.json"):
            if path.parent.name.lower() == ARCHIVE_NAME:
                continue
            try:
                report = json.loads(path.read_text(encoding="utf-8"))
            except (OSError, json.JSONDecodeError):
                continue
            if report.get("status") != "running":
                continue
            report["status"] = "failed"
            report["interrupted"] = True
            report["finished_at"] = _now()
            for scheme in report.get("schemes") or []:
                if scheme.get("status") in {"running", "analyzed"}:
                    scheme["status"] = "failed"
                    scheme["error"] = "Application stopped before completion."
            message = (
                "Aplicatia a fost repornita inainte ca aceasta rulare sa se incheie."
            )
            report.setdefault("events", []).append(
                {
                    "sequence": len(report.get("events") or []) + 1,
                    "timestamp": report["finished_at"],
                    "type": "interrupted",
                    "level": "error",
                    "scheme": "",
                    "message": message,
                    "details": {},
                }
            )
            self._append_unique_log(
                report,
                level="error",
                text=message,
                timestamp=report["finished_at"],
            )
            self._write(report)

    @staticmethod
    def _scheme(report: dict, event: dict) -> dict | None:
        index = event.get("index")
        scheme_code = str(event.get("scheme") or "")
        if index is not None:
            match = next(
                (
                    scheme
                    for scheme in report.get("schemes") or []
                    if scheme.get("index") == index
                ),
                None,
            )
            if match:
                if scheme_code and not match.get("scheme"):
                    match["scheme"] = scheme_code
                return match
        if scheme_code:
            match = next(
                (
                    scheme
                    for scheme in report.get("schemes") or []
                    if scheme.get("scheme") == scheme_code
                ),
                None,
            )
            if match:
                return match
        schemes = report.get("schemes") or []
        return schemes[0] if len(schemes) == 1 else None

    @staticmethod
    def _start_line(
        *,
        kind: str,
        system: str,
        allow_commit: bool,
        auto_classify: bool,
        schemes: list[dict],
        source_file: str,
    ) -> str:
        names = [item.get("scheme") for item in schemes if item.get("scheme")]
        mode = "salvare SAP" if allow_commit else "dry-run"
        classify = (
            "cu mentinere automata" if auto_classify else "fara mentinere automata"
        )
        source = f" fisier={source_file}." if source_file else ""
        scheme_txt = ", ".join(names) if names else "fara scheme predefinite"
        return (
            f"Start {kind or 'job'} pe {system or '-'} ({mode}, {classify}). "
            f"Scheme: {scheme_txt}.{source}"
        )

    def _log_entry(self, report: dict, event: dict, event_record: dict) -> dict | None:
        text, level = self._humanize(event, event_record)
        if not text:
            return None
        return {
            "sequence": len(report.get("log") or []) + 1,
            "timestamp": event_record.get("timestamp") or _now(),
            "level": level,
            "scheme": str(event.get("scheme") or event_record.get("scheme") or ""),
            "text": text,
        }

    def _append_unique_log(
        self,
        report: dict,
        *,
        level: str,
        text: str,
        scheme: str = "",
        timestamp: str = "",
    ) -> None:
        entries = report.setdefault("log", [])
        if entries and entries[-1].get("text") == text:
            return
        entries.append(
            {
                "sequence": len(entries) + 1,
                "timestamp": timestamp or _now(),
                "level": level,
                "scheme": scheme,
                "text": text,
            }
        )

    @staticmethod
    def _humanize(event: dict, event_record: dict) -> tuple[str, str]:
        event_type = str(event.get("type") or "")
        message = event_record.get("message") or ""
        scheme = str(event.get("scheme") or "")
        variant = str(event.get("variant") or "")
        summary = event.get("summary") if isinstance(event.get("summary"), dict) else {}
        classify = (
            event.get("classify_summary")
            if isinstance(event.get("classify_summary"), dict)
            else {}
        )
        descriptions = (
            event.get("description_summary")
            if isinstance(event.get("description_summary"), dict)
            else {}
        )
        remaining = event.get("remaining_worklist")
        ok = event.get("ok")
        status = str(event.get("status") or "")
        level = "info"
        if event_type in {"fatal", "interrupted"} or ok is False:
            level = "error"
        elif event.get("level") in {"error", "warn", "warning"}:
            level = "warn" if event.get("level") in {"warn", "warning"} else "error"
        elif status in {"no_data", "no_variant"}:
            level = "warn"

        if event_type == "queue-start":
            return f"Coada pornita ({event.get('total') or 0} scheme).", level
        if event_type == "job-start":
            return f"Clasificare pornita ({event.get('total') or 0} grupuri).", level
        if event_type == "scheme-start":
            return (
                f"Schema pornita. Varianta aplicata: {variant or '-'}."
                + (f" Schema SAP: {scheme}." if scheme else ""),
                level,
            )
        if event_type in {"scheme-analysis", "analysis"}:
            counts = _counts(summary, "ready", "blocked", "groups")
            source = event.get("source") or event.get("file") or ""
            extra = f" Sursa: {source}." if source else ""
            return f"Analiza worklist: {counts or 'fara sumar'}.{extra}", level
        if event_type == "group-start":
            return (
                f"Grup HS {event.get('hs_code') or '-'} "
                f"({event.get('count') or 0} materiale).",
                level,
            )
        if event_type == "group-end":
            return (
                f"Confirmare grup HS {event.get('hs_code') or '-'}: "
                f"{message or ('ok' if ok else 'eroare')}.",
                level,
            )
        if event_type == "material-end":
            product = event.get("product") or event.get("material") or "-"
            return (
                f"Material {product}: HS {event.get('hs_code') or '-'} "
                f"-> {status or 'unknown'}"
                + (f" ({message})" if message else "")
                + ".",
                "warn" if status in {"blocked", "dry_run"} else level,
            )
        if event_type == "description-preflight":
            return message or "Precontrol descrieri comerciale.", level
        if event_type == "description-start":
            return "Mentinere descrieri DE/EN pornita.", level
        if event_type == "description-item-end":
            product = event.get("material") or event.get("product") or "-"
            outcome = "salvate" if ok else (status or "esuate")
            return (
                f"Material {product}: descrieri {outcome}"
                + (f" ({message})" if message else "")
                + ".",
                level,
            )
        if event_type == "description-end":
            desc = descriptions or summary
            counts = _counts(desc, "saved", "total", "failed", "not_found")
            return (
                f"Descrieri DE/EN incheiate ({counts or message or 'fara sumar'}).",
                level,
            )
        if event_type == "remaining-worklist":
            artifact = remaining if isinstance(remaining, dict) else {}
            if artifact.get("ok"):
                return (
                    "Worklist ramas salvat: "
                    f"{artifact.get('file') or '-'} "
                    f"(ramase={artifact.get('remaining_products')}, "
                    f"eliminate={artifact.get('removed_products')}).",
                    "warn" if artifact.get("warning") else "info",
                )
            error = artifact.get("error") or "fisier indisponibil"
            return f"Worklist-ul ramas NU a putut fi creat: {error}.", "error"
        if event_type == "scheme-end":
            if status == "no_data":
                return (
                    "Worklist gol; schema omisa, continui cu urmatoarea."
                    + (f" {message}" if message else ""),
                    "warn",
                )
            if status == "no_variant":
                return (
                    "Varianta SAP lipsa; schema omisa, continui cu urmatoarea."
                    + (f" {message}" if message else ""),
                    "warn",
                )
            hs_txt = _counts(
                classify, "materials_committed", "materials_failed", "failed", "groups"
            )
            desc_txt = _counts(descriptions, "saved", "total", "failed", "not_found")
            if ok:
                return (
                    "Schema finalizata."
                    + (f" HS: {hs_txt}." if hs_txt else "")
                    + (f" DE/EN: {desc_txt}." if desc_txt else "")
                    + (f" {message}" if message else ""),
                    level,
                )
            return f"Schema esuata: {message or 'eroare necunoscuta'}.", "error"
        if event_type == "job-end":
            counts = _counts(
                summary, "materials_committed", "failed", "materials_failed"
            )
            if ok is False:
                return (
                    f"Job esuat: {message or event.get('error') or 'eroare'}.",
                    "error",
                )
            return f"Job finalizat ({counts or 'fara sumar'}).", level
        if event_type == "queue-end":
            counts = _counts(
                summary, "completed", "failed", "skipped", "no_data", "no_variant"
            )
            abort = _first_line((summary or {}).get("abort_reason"))
            text = f"Coada finalizata ({counts or 'fara sumar'})."
            if abort:
                text += f" Oprire: {abort}"
                level = "error"
            return text, level
        if event_type == "fatal":
            return f"Eroare fatala: {message or 'necunoscuta'}.", "error"
        if event_type == "interrupted":
            return message or "Rulare intrerupta.", "error"
        if event_type in {"log", "step"}:
            return message, level if event.get("level") else "info"
        if event_type == "download-ready":
            return (
                f"Worklist descarcat: {event.get('file') or event.get('path') or '-'}.",
                "info",
            )
        if event_type == "session-ready":
            return "Sesiune SAP gata.", "info"
        if event_type == "_end_":
            return "", "info"
        return message, level

    @staticmethod
    def _event_record(report: dict, event: dict) -> dict:
        allowed = (
            "index",
            "total",
            "scheme",
            "variant",
            "name",
            "level",
            "hs_code",
            "count",
            "status",
            "selected",
            "requested",
            "confirmed",
            "missing",
            "material",
            "product",
            "msg",
            "ok",
            "file",
            "ready",
            "products",
            "languages",
        )
        details = {key: event.get(key) for key in allowed if key in event}
        message = ""
        if event.get("type") in {"log", "step"}:
            message = _first_line(event.get("msg"))
        elif event.get("type") == "fatal":
            message = _first_line(event.get("error"))
        elif event.get("msg"):
            message = _first_line(event.get("msg"))
        return {
            "sequence": len(report.get("events") or []) + 1,
            "timestamp": _now(),
            "type": str(event.get("type") or "event"),
            "level": str(event.get("level") or ""),
            "scheme": str(event.get("scheme") or ""),
            "message": message,
            "details": details,
        }

    def _apply_event(self, report: dict, event: dict) -> None:
        event_type = event.get("type")
        scheme = self._scheme(report, event)
        if event_type == "scheme-start" and scheme:
            scheme["status"] = "running"
            scheme["variant"] = str(event.get("variant") or scheme.get("variant") or "")
        elif event_type == "job-start" and scheme:
            scheme["status"] = "running"
        elif event_type in {"scheme-analysis", "analysis"} and scheme:
            scheme["status"] = "analyzed"
            scheme["analysis"] = event.get("summary")
            scheme["analysis_id"] = event.get("analysis")
            scheme["source_file"] = str(event.get("source") or event.get("file") or "")
        elif event_type == "material-end":
            self._update_material(report, event, description=False)
        elif event_type == "description-item-end":
            self._update_material(report, event, description=True)
        elif event_type == "remaining-worklist" and scheme:
            self._set_remaining_artifact(
                report, scheme, event.get("remaining_worklist")
            )
        elif event_type == "scheme-end" and scheme:
            status = str(event.get("status") or "")
            if status in {"no_data", "no_variant"}:
                scheme["status"] = status
            else:
                scheme["status"] = "completed" if event.get("ok") else "failed"
            scheme["classification"] = event.get("classify_summary")
            scheme["descriptions"] = event.get("description_summary")
            self._set_remaining_artifact(
                report, scheme, event.get("remaining_worklist")
            )
            scheme["error"] = _first_line(event.get("msg"))
        elif event_type == "job-end" and scheme:
            scheme["status"] = "completed" if event.get("ok", True) else "failed"
            summary = event.get("summary") or {}
            scheme["classification"] = (
                None if summary.get("descriptions_only") else event.get("summary")
            )
            descriptions = summary.get("description_summary")
            scheme["descriptions"] = descriptions
            self._set_remaining_artifact(
                report, scheme, event.get("remaining_worklist")
            )
            scheme["error"] = _first_line(event.get("error"))
            report["summary"] = event.get("summary")
            report["status"] = (
                "failed"
                if not event.get("ok", True)
                else "warning"
                if report.get("traceability_errors")
                else "completed"
            )
            report["error"] = _first_line(event.get("error"))
        elif event_type == "queue-end":
            report["summary"] = event.get("summary")
            failed = int((event.get("summary") or {}).get("failed") or 0)
            report["status"] = (
                "failed"
                if failed
                else "warning"
                if report.get("traceability_errors")
                else "completed"
            )
            report["error"] = _first_line(
                (event.get("summary") or {}).get("abort_reason")
            )
        elif event_type == "fatal":
            report["status"] = "failed"
            report["error"] = _first_line(event.get("error"))

        if event_type in {"job-end", "queue-end", "fatal"}:
            report["finished_at"] = _now()

    @staticmethod
    def _set_remaining_artifact(
        report: dict, scheme: dict, artifact: dict | None
    ) -> None:
        scheme["remaining_worklist"] = artifact
        if artifact and (artifact.get("ok") is False or artifact.get("warning")):
            error = _first_line(artifact.get("error") or artifact.get("warning"))
            scheme["traceability_error"] = error
            traceability_errors = report.setdefault("traceability_errors", [])
            if error and error not in traceability_errors:
                traceability_errors.append(error)

    @staticmethod
    def _update_material(report: dict, event: dict, *, description: bool) -> None:
        product = str(event.get("product") or event.get("material") or "").strip()
        if not product:
            return
        scheme_code = str(event.get("scheme") or "")
        material = next(
            (
                item
                for item in report.get("materials") or []
                if item.get("product") == product
                and (not scheme_code or item.get("scheme") == scheme_code)
            ),
            None,
        )
        if material is None:
            material = {
                "scheme": scheme_code,
                "product": product,
                "hs_code": "",
                "hs_status": "pending",
                "hs_message": "",
                "description_status": "pending",
                "description_message": "",
            }
            report["materials"].append(material)
        if description:
            status = str(event.get("status") or "").strip().lower()
            if event.get("ok"):
                material["description_status"] = status or "confirmed"
            else:
                material["description_status"] = status or "failed"
            material["description_message"] = _first_line(event.get("msg"))
        else:
            material["hs_code"] = str(event.get("hs_code") or "")
            material["hs_status"] = str(event.get("status") or "unknown")
            material["hs_message"] = _first_line(event.get("msg"))
