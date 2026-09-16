"""Local Flask server for SAP GTS Customs Description uploader.

Endpoints:
  GET  /                  -> upload UI
  GET  /env               -> environment label + production warning flag
  GET  /config            -> current config
  POST /config            -> update config
  POST /preview           -> upload xlsx, return parsed rows
  POST /session/start     -> start persistent SAP browser session
  POST /session/stop      -> close session
  GET  /session/status    -> session state
  GET  /session/events    -> SSE stream of session-level events (login etc)
  POST /reset-browser     -> kill leftover Chromium + remove lock files
  POST /run               -> dispatch job to running session, returns job id
  POST /stop/<job>        -> request graceful cancellation of job
  GET  /events/<job>      -> SSE stream of job progress events
  GET  /report/<job>      -> CSV with results once finished
"""

from __future__ import annotations

import csv
import io
import json
import os
import queue
import re
import shutil
import sys
import threading
import time
import uuid
from pathlib import Path
from urllib.parse import urlsplit

# Ensure sibling modules (excel_reader, sap_automation) are importable
# regardless of the working directory or Python distribution used.
_SERVER_DIR = str(Path(__file__).resolve().parent)
if _SERVER_DIR not in sys.path:
    sys.path.insert(0, _SERVER_DIR)

from excel_reader import read_materials
from flask import Flask, Response, jsonify, render_template, request, send_file
from hs_automation import (
    cleanup_hs_lock,
    commercial_description_url,
    get_hs_session,
)
from hs_reference import RPA_SCHEME_ORDER, HsReference, list_schemes, scheme_info
from hs_reporting import HsReportStore
from hs_worklist import (
    REASON_LABELS,
    analyze,
    blocked_workbook_rows,
    build_description_plan,
    read_worklist,
    save_remaining_worklist,
    selected_groups,
)
from sap_automation import cleanup_user_data_lock, get_session

ROOT = Path(__file__).resolve().parent.parent
UPLOADS = ROOT / "uploads"
UPLOADS.mkdir(exist_ok=True)
CONFIG_PATH = ROOT / "config.json"
HS_REPORTS_DIR = ROOT / "reports-hs"

# openpyxl does not read the legacy binary .xls format. Advertising it as
# supported only postpones the error until parsing, so reject it at upload.
ALLOWED_UPLOAD_SUFFIXES = {".xlsx", ".xlsm"}
MAX_UPLOAD_BYTES = 64 * 1024 * 1024

app = Flask(
    __name__,
    template_folder=str(Path(__file__).parent / "templates"),
    static_folder=str(Path(__file__).parent / "static"),
    static_url_path="/static",
)

# job queues
JOBS: dict[str, queue.Queue] = {}
JOB_RESULTS: dict[str, dict] = {}
JOB_CANCEL: dict[str, threading.Event] = {}

# session-level event queue (login progress, etc)
SESSION_Q: queue.Queue = queue.Queue()


def session_event(ev: dict) -> None:
    SESSION_Q.put(ev)


def load_config() -> dict:
    return json.loads(CONFIG_PATH.read_text(encoding="utf-8-sig"))


def save_config(cfg: dict) -> None:
    CONFIG_PATH.write_text(
        json.dumps(cfg, indent=2, ensure_ascii=False), encoding="utf-8"
    )


def is_production(cfg: dict) -> bool:
    url = (cfg.get("sap_url") or "").lower()
    for host in cfg.get("production_hosts") or []:
        if host.lower() in url:
            return True
    return False


def system_label_for(cfg: dict, url: str) -> str:
    """Name the SAP system behind a URL so the badge cannot contradict the target."""
    host = (urlsplit(str(url or "")).hostname or "").lower()
    if host:
        for system in (cfg.get("hs") or {}).get("systems") or []:
            candidate = (
                urlsplit(str(system.get("sap_url") or "")).hostname or ""
            ).lower()
            if candidate and candidate == host:
                label = str(system.get("id") or "").strip()
                if label:
                    return label
    return str(cfg.get("environment_label") or "").strip()


def language_selection_error(cfg: dict) -> str | None:
    enabled = [
        language for language in cfg.get("languages") or [] if language.get("enabled")
    ]
    if not enabled:
        return "Select at least one language (DE or EN)."
    return None


def safe_upload_target(filename: str, folder: Path) -> Path:
    """Reject path traversal and unexpected file types on upload."""
    name = Path(filename or "").name
    stem = re.sub(r"[^A-Za-z0-9._-]+", "_", Path(name).stem).strip("._") or "upload"
    suffix = Path(name).suffix.lower()
    if suffix not in ALLOWED_UPLOAD_SUFFIXES:
        raise ValueError(
            "Sunt acceptate doar fisiere Excel ("
            + ", ".join(sorted(ALLOWED_UPLOAD_SUFFIXES))
            + ")."
        )
    folder.mkdir(parents=True, exist_ok=True)
    return folder / f"{stem}{suffix}"


def store_upload(file_storage, folder: Path) -> Path:
    target = safe_upload_target(file_storage.filename, folder)
    file_storage.save(str(target))
    if target.stat().st_size > MAX_UPLOAD_BYTES:
        target.unlink(missing_ok=True)
        raise ValueError("Fisierul depaseste limita de 64 MB.")
    return target


# ---------- basic info ----------


def _no_store(template: str):
    resp = app.make_response(render_template(template))
    resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    resp.headers["Pragma"] = "no-cache"
    resp.headers["Expires"] = "0"
    return resp


@app.get("/")
def launcher():
    return _no_store("launcher.html")


@app.get("/apps/descriptions")
def app_descriptions():
    return _no_store("index.html")


@app.get("/apps/hs")
def app_hs():
    return _no_store("hs_codes.html")


@app.get("/apps/hs/reports")
def app_hs_reports():
    return _no_store("hs_reports.html")


@app.get("/apps/hs/dashboard")
def app_hs_dashboard():
    return _no_store("hs_dashboard.html")


@app.get("/env")
def env():
    cfg = load_config()
    hs_url = (cfg.get("hs") or {}).get("sap_url", "")
    return jsonify(
        {
            "label": system_label_for(cfg, cfg.get("sap_url", "")),
            "url": cfg.get("sap_url", ""),
            "hs_label": system_label_for(cfg, hs_url),
            "hs_url": hs_url,
            "is_production": is_production(cfg),
        }
    )


@app.get("/config")
def get_config():
    return jsonify(load_config())


@app.post("/config")
def post_config():
    cfg = load_config()
    incoming = request.get_json(force=True) or {}
    cfg.update(incoming)
    if err := language_selection_error(cfg):
        return jsonify({"error": err}), 400
    save_config(cfg)
    return jsonify(cfg)


@app.post("/preview")
def preview():
    f = request.files.get("file")
    if not f or not f.filename:
        return jsonify({"error": "no file"}), 400
    cfg = load_config()
    if err := language_selection_error(cfg):
        return jsonify({"error": err}), 400
    excel_cfg = cfg.get("excel", {})
    languages = cfg.get("languages") or [
        {
            "code": "DE",
            "column": excel_cfg.get("description_column", "J"),
            "enabled": True,
        }
    ]
    try:
        target = store_upload(f, UPLOADS)
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    items = read_materials(
        str(target),
        material_column=excel_cfg.get("material_column", "B"),
        languages=languages,
        header_row=int(excel_cfg.get("header_row", 1)),
        sheet=excel_cfg.get("sheet"),
    )
    return jsonify(
        {
            "file": f.filename,
            "count": len(items),
            "languages": [
                {
                    "code": lng["code"],
                    "column": lng.get("column"),
                    "enabled": lng.get("enabled", True),
                }
                for lng in languages
            ],
            "preview": items[:50],
            "all": items,
        }
    )


# ---------- SAP session lifecycle ----------


@app.post("/session/start")
def session_start():
    cfg = load_config()
    payload = request.get_json(silent=True) or {}
    creds = None
    if payload.get("username") and payload.get("password"):
        creds = {"username": payload["username"], "password": payload["password"]}
    if "headless" in payload:
        cfg["headless"] = bool(payload.get("headless"))
    sess = get_session()
    res = sess.start(cfg, creds, session_event)
    return jsonify(res)


@app.post("/session/stop")
def session_stop():
    sess = get_session()
    sess.stop()
    return jsonify({"ok": True})


@app.get("/session/status")
def session_status():
    return jsonify(get_session().state)


@app.get("/session/events")
def session_events():
    def stream():
        # initial state ping
        yield f"data: {json.dumps({'type': 'state', **get_session().state})}\n\n"
        while True:
            ev = SESSION_Q.get()
            yield f"data: {json.dumps(ev, ensure_ascii=False)}\n\n"

    return Response(stream(), mimetype="text/event-stream")


@app.post("/reset-browser")
def reset_browser():
    sess = get_session()
    if sess.state["state"] in ("ready", "busy", "starting"):
        sess.stop()
    rep = cleanup_user_data_lock(kill_chrome=True)
    return jsonify(rep)


# ---------- Job execution ----------


@app.post("/run")
def run():
    payload = request.get_json(force=True) or {}
    items = payload.get("items") or []
    limit = int(payload.get("limit") or 0)
    if not items:
        return jsonify({"error": "no items"}), 400
    if limit > 0:
        items = items[:limit]

    cfg = load_config()
    if err := language_selection_error(cfg):
        return jsonify({"error": err}), 400
    if "headless" in payload:
        cfg["headless"] = bool(payload.get("headless"))
    creds = None
    cred_in = payload.get("credentials") or {}
    if cred_in.get("username") and cred_in.get("password"):
        creds = {
            "username": cred_in["username"],
            "password": cred_in["password"],
        }

    sess = get_session()
    # Auto-start the session if it's not already running.
    # The job is queued anyway and will execute once the worker is ready.
    if sess.state["state"] in ("stopped", "error"):
        sess.start(cfg, creds, session_event)

    job_id = uuid.uuid4().hex[:8]
    q: queue.Queue = queue.Queue()
    cancel_evt = threading.Event()
    JOBS[job_id] = q
    JOB_CANCEL[job_id] = cancel_evt

    job_result: dict = {"summary": None, "items": []}
    JOB_RESULTS[job_id] = job_result

    def progress(ev):
        event_type = ev.get("type")
        if event_type == "job-end":
            job_result["summary"] = ev.get("summary")
        if event_type == "item-end":
            job_result["items"].append(
                {
                    "material": ev.get("material"),
                    "ok": ev.get("ok"),
                    "msg": ev.get("msg"),
                    "not_found": ev.get("not_found", False),
                }
            )
        if event_type == "fatal":
            completed = job_result["items"]
            not_found = sum(1 for item in completed if item.get("not_found"))
            ok = sum(1 for item in completed if item.get("ok"))
            completed_fail = len(completed) - ok - not_found
            remaining = max(0, len(items) - len(completed))
            error = ev.get("error", "")
            job_result["summary"] = {
                "total": len(items),
                "ok": ok,
                "fail": completed_fail + (1 if remaining else 0),
                "not_found": not_found,
                "skipped": max(0, remaining - 1),
                "error": error,
            }
            job_result["items"].append(
                {
                    "material": "(job)",
                    "ok": False,
                    "msg": f"Job failed before completion: {error}",
                    "not_found": False,
                }
            )
        q.put(ev)

    sess.run_items(items, cfg, event_cb=progress, cancel_event=cancel_evt)
    return jsonify({"job": job_id, "count": len(items)})


@app.post("/stop/<job>")
def stop(job):
    evt = JOB_CANCEL.get(job)
    if not evt:
        return jsonify({"error": "unknown job"}), 404
    evt.set()
    return jsonify({"status": "cancel requested"})


@app.get("/events/<job>")
def events(job):
    q = JOBS.get(job)
    if not q:
        return jsonify({"error": "unknown job"}), 404

    def stream():
        while True:
            ev = q.get()
            if ev.get("type") == "_end_":
                yield "event: end\ndata: {}\n\n"
                break
            yield f"data: {json.dumps(ev, ensure_ascii=False)}\n\n"

    return Response(stream(), mimetype="text/event-stream")


@app.get("/report/<job>")
def report(job):
    res = JOB_RESULTS.get(job)
    if not res or res.get("summary") is None:
        return jsonify({"error": "report not ready"}), 404
    buf = io.StringIO()
    w = csv.writer(buf, delimiter=";")
    w.writerow(["material", "status", "message"])
    for d in res.get("items", []):
        status = (
            "OK" if d.get("ok") else ("NOT_FOUND" if d.get("not_found") else "FAIL")
        )
        w.writerow([d.get("material", ""), status, d.get("msg", "")])
    return Response(
        buf.getvalue(),
        mimetype="text/csv",
        headers={"Content-Disposition": f"attachment; filename=report_{job}.csv"},
    )


# ---------- HS / tariff classification app ----------

HS_UPLOADS = ROOT / "uploads-hs"
HS_REFERENCE_DIR = ROOT / "Referinta"
HS_REFERENCE_NAME = "Referinta TECDOC HS.xlsx"
HS_ANALYSES: dict[str, dict] = {}
HS_JOBS: dict[str, queue.Queue] = {}
HS_JOB_CANCEL: dict[str, threading.Event] = {}
HS_SESSION_Q: queue.Queue = queue.Queue()
HS_REPORT_STORE = HsReportStore(HS_REPORTS_DIR)
_REFERENCE_CACHE_LOCK = threading.Lock()
_REFERENCE_CACHE: dict = {}


def hs_session_event(ev: dict) -> None:
    HS_SESSION_Q.put(ev)


DEFAULT_HS_SYSTEMS = [
    {
        "id": "FT1",
        "label": "FT1 (cift1)",
        "sap_url": "https://cift1.dc.hella.com:44300/sap/bc/ui2/flp"
        "?sap-client=100&sap-language=EN"
        "#CustomsProduct-classify?sap-ui-tech-hint=GUI",
    },
    {
        "id": "FT6",
        "label": "FT6 (cift6)",
        "sap_url": "https://cift6.dc.hella.com:44300/sap/bc/ui2/flp"
        "?sap-client=100&sap-language=EN"
        "#CustomsProduct-classify?sap-ui-tech-hint=GUI",
    },
]


def hs_config(cfg: dict | None = None) -> dict:
    cfg = cfg or load_config()
    hs = dict(cfg.get("hs") or {})
    hs.setdefault("sap_url", cfg.get("sap_url", ""))
    hs.setdefault("reference_file", f"{HS_REFERENCE_DIR.name}/{HS_REFERENCE_NAME}")
    hs.setdefault("default_scheme", "EDCHSCDEEX")
    hs.setdefault("allow_commit", False)
    hs["allow_commit"] = hs.get("allow_commit") is True
    hs.setdefault("display_all_products", False)
    hs["display_all_products"] = hs.get("display_all_products") is True
    hs.setdefault(
        "description_display_maintained",
        cfg.get("display_maintained", True),
    )
    hs["description_display_maintained"] = (
        hs.get("description_display_maintained") is not False
    )
    configured_languages = hs.get("description_languages") or cfg.get("languages") or []
    by_code = {
        str(language.get("code") or "").strip().upper(): language
        for language in configured_languages
        if language.get("code")
    }
    defaults = {
        "DE": {"code": "DE", "country_name": "Germany", "column": "J"},
        "EN": {
            "code": "EN",
            "country_name": "United Kingdom",
            "column": "K",
        },
    }
    # DE/EN columns stay configured for the integrated HS queue. Which language
    # is written is decided later from the worklist Country/Country Group
    # (EU -> DE, GB -> EN). The old `enabled` flag still controls only the
    # standalone descriptions screen.
    hs["description_languages"] = [
        {
            **defaults[code],
            **{
                key: value
                for key, value in (by_code.get(code) or {}).items()
                if key in ("code", "country_name", "column")
            },
            "code": code,
        }
        for code in ("DE", "EN")
    ]
    systems = [s for s in (hs.get("systems") or []) if s.get("id") and s.get("sap_url")]
    hs["systems"] = systems or list(DEFAULT_HS_SYSTEMS)
    return hs


def hs_active_system(hs: dict) -> str:
    url = (hs.get("sap_url") or "").strip()
    for system in hs.get("systems") or []:
        if system["sap_url"].strip() == url:
            return system["id"]
    return ""


def load_reference(hs_cfg: dict) -> HsReference:
    raw = hs_cfg.get("reference_file") or ""
    path = Path(raw)
    if not path.is_absolute():
        path = ROOT / path
    if not path.exists():
        raise FileNotFoundError(f"Fisierul de referinta lipseste: {path}")
    resolved = path.resolve()
    stat = resolved.stat()
    key = (
        str(resolved),
        hs_cfg.get("reference_sheet"),
        stat.st_mtime_ns,
        stat.st_size,
    )
    with _REFERENCE_CACHE_LOCK:
        if _REFERENCE_CACHE.get("key") == key:
            return _REFERENCE_CACHE["reference"]
    reference = HsReference(resolved, hs_cfg.get("reference_sheet"))
    with _REFERENCE_CACHE_LOCK:
        _REFERENCE_CACHE.clear()
        _REFERENCE_CACHE.update({"key": key, "reference": reference})
    return reference


def reference_status(hs_cfg: dict) -> dict:
    """Load/refresh the reference and expose its startup health to the UI."""
    try:
        reference = load_reference(hs_cfg)
        reference.validate_description_columns(hs_cfg["description_languages"])
        return {
            "reference_exists": True,
            "reference_loaded": True,
            "reference_error": "",
            "reference_tecdoc_count": reference.tecdoc_count,
            "reference_header_row": reference.header_row,
            "description_languages": hs_cfg["description_languages"],
        }
    except Exception as exc:
        raw = Path(hs_cfg.get("reference_file") or "")
        if not raw.is_absolute():
            raw = ROOT / raw
        return {
            "reference_exists": raw.exists(),
            "reference_loaded": False,
            "reference_error": str(exc),
            "reference_tecdoc_count": 0,
            "reference_header_row": 0,
            "description_languages": hs_cfg["description_languages"],
        }


# Prime the mtime-aware cache while the server imports. A missing/invalid file
# is reported in /hs/config instead of preventing the local UI from starting.
REFERENCE_STARTUP_STATUS = reference_status(hs_config())


def store_analysis(
    result: dict,
    source: str,
    source_path: str | Path | None = None,
) -> str:
    analysis_id = uuid.uuid4().hex[:8]
    result["source"] = source
    result["_source_path"] = str(source_path or "")
    HS_ANALYSES[analysis_id] = result
    return analysis_id


def analysis_response(analysis_id: str, result: dict) -> dict:
    hs = hs_config()
    description_languages = hs.get("description_languages") or []
    try:
        reference = load_reference(hs)
    except Exception:
        reference = None

    def group_descriptions(group: dict) -> list[dict]:
        """Texts the operator approves, so the preview shows what SAP will write."""
        if reference is None:
            return []
        entries = []
        for tecdoc in group.get("tecdocs") or []:
            texts = reference.commercial_descriptions(tecdoc, description_languages)
            if texts:
                entries.append({"tecdoc": tecdoc, "texts": texts})
        return entries

    return {
        "analysis": analysis_id,
        "scheme": result["scheme"],
        "scheme_label": result["scheme_label"],
        "reference_column": result["reference_column"],
        "variant": result["variant"],
        "source": result.get("source", ""),
        "summary": result["summary"],
        "reason_labels": REASON_LABELS,
        "description_languages": [
            str(language.get("code") or "").strip().upper()
            for language in description_languages
            if language.get("code")
        ],
        "groups": [
            {
                "hs_code": g["hs_code"],
                "count": g["count"],
                "tecdocs": g["tecdocs"],
                "sample": [p["product"] for p in g["products"][:5]],
                "text": g["products"][0].get("tecdoc_text", "")
                if g["products"]
                else "",
                "descriptions": group_descriptions(g),
                "products": g["products"],
            }
            for g in result["groups"]
        ],
        "blocked_preview": result["blocked"][:100],
    }


@app.get("/hs/schemes")
def hs_schemes():
    presets = [
        {"scheme": code, "variant": scheme_info(code)["variant"]}
        for code in RPA_SCHEME_ORDER
    ]
    return jsonify(
        {
            "schemes": list_schemes(),
            "default": hs_config()["default_scheme"],
            "queue_preset": presets,
        }
    )


@app.get("/hs/config")
def hs_get_config():
    cfg = load_config()
    hs = hs_config(cfg)
    hs_target_cfg = {**cfg, "sap_url": hs.get("sap_url", "")}
    return jsonify(
        {
            **hs,
            "system": hs_active_system(hs),
            "environment_label": cfg.get("environment_label", ""),
            "is_production": is_production(hs_target_cfg),
            **reference_status(hs),
        }
    )


@app.post("/hs/config")
def hs_post_config():
    cfg = load_config()
    incoming = request.get_json(force=True) or {}
    hs = hs_config(cfg)

    # The SAP target may only be switched between configured systems.
    if "system" in incoming:
        wanted = str(incoming.get("system") or "").strip()
        match = next((s for s in hs["systems"] if s["id"] == wanted), None)
        if not match:
            return jsonify({"error": f"Sistem necunoscut: {wanted!r}"}), 400
        sess = get_hs_session()
        if sess.state["state"] not in ("stopped", "error"):
            return jsonify(
                {
                    "error": "Opreste sesiunea SAP inainte de a schimba sistemul "
                    f"(state={sess.state['state']})."
                }
            ), 409
        hs["sap_url"] = match["sap_url"]

    for key in ("reference_file", "reference_sheet"):
        if key in incoming:
            hs[key] = incoming[key]
    if "default_scheme" in incoming:
        default_scheme = str(incoming["default_scheme"] or "").strip()
        try:
            scheme_info(default_scheme)
        except KeyError:
            return jsonify({"error": "Schema implicita nu este configurata."}), 400
        hs["default_scheme"] = default_scheme
    if "description_languages" in incoming:
        raw_languages = incoming["description_languages"]
        if not isinstance(raw_languages, list):
            return jsonify(
                {"error": "description_languages trebuie sa fie lista."}
            ), 400
        normalized_languages = []
        for raw_language in raw_languages:
            if not isinstance(raw_language, dict):
                return jsonify(
                    {"error": "Fiecare limba trebuie sa fie un obiect."}
                ), 400
            code = str(raw_language.get("code") or "").strip().upper()
            column = str(raw_language.get("column") or "").strip()
            country_name = str(raw_language.get("country_name") or "").strip()
            if code not in {"DE", "EN"} or not column or not country_name:
                return jsonify(
                    {
                        "error": "Maparea integrata necesita cod, coloana si tara "
                        "pentru DE si EN."
                    }
                ), 400
            normalized_languages.append(
                {"code": code, "column": column, "country_name": country_name}
            )
        if {item["code"] for item in normalized_languages} != {"DE", "EN"} or len(
            normalized_languages
        ) != 2:
            return jsonify(
                {"error": "Maparea integrata trebuie sa contina exact DE si EN."}
            ), 400
        hs["description_languages"] = normalized_languages
    if "description_display_maintained" in incoming:
        if not isinstance(incoming["description_display_maintained"], bool):
            return jsonify(
                {"error": "description_display_maintained trebuie sa fie boolean JSON."}
            ), 400
        hs["description_display_maintained"] = incoming[
            "description_display_maintained"
        ]
    if "allow_commit" in incoming:
        if not isinstance(incoming["allow_commit"], bool):
            return jsonify({"error": "allow_commit trebuie sa fie boolean JSON."}), 400
        hs["allow_commit"] = incoming["allow_commit"] is True
    if "display_all_products" in incoming:
        if not isinstance(incoming["display_all_products"], bool):
            return jsonify(
                {"error": "display_all_products trebuie sa fie boolean JSON."}
            ), 400
        hs["display_all_products"] = incoming["display_all_products"] is True
    cfg["hs"] = hs
    save_config(cfg)
    if "reference_file" in incoming or "reference_sheet" in incoming:
        with _REFERENCE_CACHE_LOCK:
            _REFERENCE_CACHE.clear()
    return jsonify({**hs, "system": hs_active_system(hs)})


@app.post("/hs/reference")
def hs_upload_reference():
    f = request.files.get("file")
    if not f or not f.filename:
        return jsonify({"error": "no file"}), 400
    try:
        staged = store_upload(f, HS_UPLOADS)
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    try:
        HsReference(staged)
    except Exception as exc:
        return jsonify({"error": f"Fisier de referinta invalid: {exc}"}), 400
    HS_REFERENCE_DIR.mkdir(parents=True, exist_ok=True)
    target = HS_REFERENCE_DIR / HS_REFERENCE_NAME
    shutil.copyfile(staged, target)
    staged.unlink(missing_ok=True)
    cfg = load_config()
    hs = hs_config(cfg)
    hs["reference_file"] = str(target.relative_to(ROOT))
    cfg["hs"] = hs
    save_config(cfg)
    with _REFERENCE_CACHE_LOCK:
        _REFERENCE_CACHE.clear()
    loaded = load_reference(hs_config(cfg))
    return jsonify(
        {
            "file": target.name,
            "tecdoc_count": loaded.tecdoc_count,
            "columns": [h for h in loaded.headers if h],
            "description_languages": hs_config(cfg)["description_languages"],
        }
    )


@app.post("/hs/analyze")
def hs_analyze():
    """Semi-automatic path: the operator uploads a worklist they exported."""
    f = request.files.get("file")
    scheme = (request.form.get("scheme") or "").strip()
    if not f or not f.filename:
        return jsonify({"error": "no file"}), 400
    if not scheme:
        return jsonify({"error": "Selecteaza schema de numerotare incarcata."}), 400
    try:
        target = store_upload(f, HS_UPLOADS)
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400

    hs_cfg = hs_config()
    try:
        reference = load_reference(hs_cfg)
        worklist = read_worklist(target)
        result = analyze(worklist, reference, scheme)
    except (KeyError, ValueError, FileNotFoundError) as exc:
        return jsonify({"error": str(exc)}), 400

    analysis_id = store_analysis(result, target.name, target)
    return jsonify(analysis_response(analysis_id, result))


@app.get("/hs/blocked/<analysis>")
def hs_blocked(analysis):
    result = HS_ANALYSES.get(analysis)
    if not result:
        return jsonify({"error": "unknown analysis"}), 404
    buf = io.StringIO()
    writer = csv.writer(buf, delimiter=";")
    for row in blocked_workbook_rows(result["blocked"]):
        writer.writerow(row)
    return Response(
        buf.getvalue(),
        mimetype="text/csv",
        headers={
            "Content-Disposition": f"attachment; filename=not_ready_{analysis}.csv"
        },
    )


@app.get("/hs/ready/<analysis>")
def hs_ready(analysis):
    result = HS_ANALYSES.get(analysis)
    if not result:
        return jsonify({"error": "unknown analysis"}), 404
    buf = io.StringIO()
    writer = csv.writer(buf, delimiter=";")
    writer.writerow(
        ["product", "product_text", "material_type", "tecdoc", "scheme", "hs_code"]
    )
    for item in result["ready"]:
        writer.writerow(
            [
                item.get("product", ""),
                item.get("text", ""),
                item.get("material_type", ""),
                item.get("tecdoc", ""),
                item.get("scheme", ""),
                item.get("hs_code", ""),
            ]
        )
    return Response(
        buf.getvalue(),
        mimetype="text/csv",
        headers={"Content-Disposition": f"attachment; filename=ready_{analysis}.csv"},
    )


def _remaining_worklist_artifact(
    report_id: str,
    *,
    scheme: str,
    index: int,
    source_path: str | Path,
    classify_summary: dict | None,
) -> dict:
    """Create the per-scheme SAP worklist minus confirmed classifications."""
    source = Path(source_path)
    if not source.is_file():
        return {
            "ok": False,
            "file": "",
            "url": "",
            "error": f"Worklistul sursa nu mai exista: {source.name or source}",
        }
    suffix = (
        source.suffix.lower()
        if source.suffix.lower() in {".xlsx", ".xlsm"}
        else ".xlsx"
    )
    filename = f"remaining_{index:02d}_{scheme}_{report_id}{suffix}"
    target = HS_REPORT_STORE.artifact_target(report_id, filename)
    committed_products = (classify_summary or {}).get("committed_products") or []
    try:
        stats = save_remaining_worklist(source, target, committed_products)
    except Exception as exc:
        return {
            "ok": False,
            "file": "",
            "url": "",
            "error": str(exc).splitlines()[0][:300],
        }
    unmatched = stats.get("unmatched_products") or []
    warning = ""
    if unmatched:
        warning = (
            f"{len(unmatched)} materiale confirmate de SAP nu au fost gasite "
            "in worklistul sursa."
        )
    return {
        "ok": True,
        **stats,
        "warning": warning,
        "url": f"/hs/reports/{report_id}/artifacts/{target.name}",
    }


@app.get("/hs/reports")
def hs_reports_list():
    limit = request.args.get("limit", 100, type=int) or 100
    return jsonify({"reports": HS_REPORT_STORE.list(limit=limit)})


@app.get("/hs/reports/stats")
def hs_reports_stats():
    return jsonify({"months": HS_REPORT_STORE.stats()})


@app.get("/hs/reports/<report_id>")
def hs_report_detail(report_id):
    report_data = HS_REPORT_STORE.get(report_id)
    if report_data is None:
        return jsonify({"error": "unknown report"}), 404
    return jsonify(report_data)


@app.delete("/hs/reports/<report_id>")
def hs_report_delete(report_id):
    if not _is_local_request():
        return (
            jsonify(
                {
                    "ok": False,
                    "error": "Stergerea rapoartelor este permisa doar de pe acest calculator.",
                }
            ),
            403,
        )
    try:
        deleted = HS_REPORT_STORE.delete(report_id)
    except ValueError:
        return jsonify({"ok": False, "error": "unknown report"}), 404
    if not deleted:
        return jsonify({"ok": False, "error": "unknown report"}), 404
    return jsonify({"ok": True, "deleted": report_id})


@app.delete("/hs/reports")
def hs_reports_clear():
    if not _is_local_request():
        return (
            jsonify(
                {
                    "ok": False,
                    "error": "Stergerea rapoartelor este permisa doar de pe acest calculator.",
                }
            ),
            403,
        )
    deleted = HS_REPORT_STORE.clear()
    return jsonify({"ok": True, "deleted": deleted})


@app.get("/hs/reports/<report_id>/artifacts/<path:filename>")
def hs_report_artifact(report_id, filename):
    path = HS_REPORT_STORE.artifact_path(report_id, filename)
    if path is None:
        return jsonify({"error": "unknown artifact"}), 404
    return send_file(path, as_attachment=True, download_name=path.name)


@app.get("/hs/reports/<report_id>/log.txt")
def hs_report_log(report_id):
    path = HS_REPORT_STORE.log_path(report_id)
    if path is None:
        report_data = HS_REPORT_STORE.get(report_id)
        if report_data is None:
            return jsonify({"error": "unknown report"}), 404
        lines = []
        for entry in report_data.get("log") or []:
            scheme = str(entry.get("scheme") or "").strip()
            prefix = f"[{scheme}] " if scheme else ""
            stamp = str(entry.get("timestamp") or "").replace("T", " ")[:19]
            lines.append(
                f"{stamp}  {str(entry.get('level') or 'info').upper():<5}  "
                f"{prefix}{entry.get('text') or ''}".rstrip()
            )
        payload = "\n".join(lines)
        if payload:
            payload += "\n"
        return Response(
            payload,
            mimetype="text/plain; charset=utf-8",
            headers={
                "Content-Disposition": f"attachment; filename=hs_report_{report_id}.log"
            },
        )
    return send_file(
        path, as_attachment=True, download_name=f"hs_report_{report_id}.log"
    )


@app.get("/hs/reports/<report_id>/materials.csv")
def hs_report_materials_csv(report_id):
    report_data = HS_REPORT_STORE.get(report_id)
    if report_data is None:
        return jsonify({"error": "unknown report"}), 404
    buf = io.StringIO()
    writer = csv.writer(buf, delimiter=";")
    writer.writerow(
        [
            "scheme",
            "product",
            "hs_code",
            "hs_status",
            "hs_message",
            "description_status",
            "description_message",
        ]
    )
    for item in report_data.get("materials") or []:
        writer.writerow(
            [
                item.get("scheme", ""),
                item.get("product", ""),
                item.get("hs_code", ""),
                item.get("hs_status", ""),
                item.get("hs_message", ""),
                item.get("description_status", ""),
                item.get("description_message", ""),
            ]
        )
    return Response(
        buf.getvalue(),
        mimetype="text/csv",
        headers={
            "Content-Disposition": f"attachment; filename=hs_report_{report_id}.csv"
        },
    )


# ---------- HS SAP session ----------


@app.post("/hs/session/start")
def hs_session_start():
    cfg = load_config()
    hs = hs_config(cfg)
    payload = request.get_json(silent=True) or {}
    creds = None
    if payload.get("username") and payload.get("password"):
        creds = {"username": payload["username"], "password": payload["password"]}
    session_cfg = {
        "sap_url": hs["sap_url"],
        "headless": bool(payload.get("headless", False)),
    }
    return jsonify(get_hs_session().start(session_cfg, creds, hs_session_event))


@app.post("/hs/session/stop")
def hs_session_stop():
    sess = get_hs_session()
    sess.stop()
    cleanup = cleanup_hs_lock(kill_chrome=True)
    return jsonify({"ok": True, **cleanup})


@app.get("/hs/session/status")
def hs_session_status():
    return jsonify(get_hs_session().state)


@app.get("/hs/session/events")
def hs_session_events():
    def stream():
        yield f"data: {json.dumps({'type': 'state', **get_hs_session().state})}\n\n"
        while True:
            ev = HS_SESSION_Q.get()
            yield f"data: {json.dumps(ev, ensure_ascii=False)}\n\n"

    return Response(stream(), mimetype="text/event-stream")


@app.post("/hs/reset-browser")
def hs_reset_browser():
    sess = get_hs_session()
    if sess.state["state"] in ("ready", "busy", "starting"):
        sess.stop()
    return jsonify(cleanup_hs_lock(kill_chrome=True))


def _hs_new_job(
    *,
    kind: str,
    hs: dict,
    items: list[dict],
    allow_commit: bool,
    auto_classify: bool,
    source_file: str = "",
) -> tuple[str, queue.Queue, threading.Event]:
    job_id = uuid.uuid4().hex[:8]
    q: queue.Queue = queue.Queue()
    cancel = threading.Event()
    HS_JOBS[job_id] = q
    HS_JOB_CANCEL[job_id] = cancel
    HS_REPORT_STORE.start(
        job_id,
        kind=kind,
        system=hs_active_system(hs) or hs.get("environment_label", ""),
        items=items,
        allow_commit=allow_commit,
        auto_classify=auto_classify,
        source_file=source_file,
    )
    return job_id, q, cancel


def _hs_record_and_queue(job_id: str, q: queue.Queue, event: dict) -> None:
    """Persist a sanitized event without allowing reporting to break SAP RPA."""
    try:
        if event.get("type") == "_end_":
            HS_REPORT_STORE.finish(job_id)
        else:
            HS_REPORT_STORE.record(job_id, event)
    except Exception as exc:
        q.put(
            {
                "type": "log",
                "level": "error",
                "msg": "Raportul persistent nu a putut fi actualizat: "
                + str(exc).splitlines()[0][:200],
            }
        )
    q.put(event)


@app.post("/hs/download")
def hs_download():
    """Full-automatic path: drive SAP to apply the variant and export."""
    payload = request.get_json(force=True) or {}
    scheme = (payload.get("scheme") or "").strip()
    try:
        info = scheme_info(scheme)
    except KeyError as exc:
        return jsonify({"error": str(exc)}), 400
    variant = (payload.get("variant") or info["variant"]).strip()

    cfg = load_config()
    hs = hs_config(cfg)
    creds = None
    cred_in = payload.get("credentials") or {}
    if cred_in.get("username") and cred_in.get("password"):
        creds = {"username": cred_in["username"], "password": cred_in["password"]}

    sess = get_hs_session()
    if sess.state["state"] in ("stopped", "error"):
        sess.start(
            {"sap_url": hs["sap_url"], "headless": bool(payload.get("headless"))},
            creds,
            hs_session_event,
        )

    job_id, q, cancel = _hs_new_job(
        kind="download",
        hs=hs,
        items=[{"scheme": info["code"], "variant": variant}],
        allow_commit=False,
        auto_classify=False,
    )

    def progress(ev):
        if ev.get("type") == "download-ready":
            _hs_record_and_queue(job_id, q, ev)
            try:
                reference = load_reference(hs)
                worklist = read_worklist(ev["path"])
                result = analyze(worklist, reference, scheme)
                analysis_id = store_analysis(
                    result,
                    ev.get("file", ""),
                    ev.get("path"),
                )
                analysis_event = {
                    "type": "analysis",
                    "index": 1,
                    **analysis_response(analysis_id, result),
                }
                _hs_record_and_queue(job_id, q, analysis_event)
                artifact = _remaining_worklist_artifact(
                    job_id,
                    scheme=info["code"],
                    index=1,
                    source_path=ev["path"],
                    classify_summary=None,
                )
                _hs_record_and_queue(
                    job_id,
                    q,
                    {
                        "type": "remaining-worklist",
                        "index": 1,
                        "scheme": info["code"],
                        "remaining_worklist": artifact,
                    },
                )
            except Exception as exc:
                _hs_record_and_queue(
                    job_id,
                    q,
                    {
                        "type": "fatal",
                        "error": f"Analiza worklistului a esuat: {exc}",
                    },
                )
            return
        _hs_record_and_queue(job_id, q, ev)

    sess.submit(
        "download",
        {
            "variant": variant,
            "scheme": scheme,
            "display_all_products": hs.get("display_all_products") is True,
        },
        progress,
        cancel,
    )
    return jsonify(
        {
            "job": job_id,
            "variant": variant,
            "scheme": info["code"],
            "report_url": f"/apps/hs/reports?report={job_id}",
        }
    )


@app.post("/hs/classify")
def hs_classify():
    """Maintain the approved HS groups. Commit is refused unless enabled."""
    payload = request.get_json(force=True) or {}
    analysis_id = str(payload.get("analysis") or "")
    result = HS_ANALYSES.get(analysis_id)
    if not result:
        return jsonify({"error": "unknown analysis"}), 404

    hs_codes = payload.get("hs_codes")
    if not isinstance(hs_codes, list):
        return jsonify(
            {"error": "Trimite explicit lista grupurilor HS selectate."}
        ), 400
    groups = selected_groups(result, hs_codes)
    if not groups:
        return jsonify({"error": "Niciun grup selectat pentru mentinere."}), 400

    hs = hs_config()
    requested_commit = payload.get("allow_commit") is True
    allow_commit = requested_commit and hs.get("allow_commit") is True
    description_plan = None
    if allow_commit:
        try:
            reference = load_reference(hs)
            description_plan = build_description_plan(
                groups, reference, hs["description_languages"]
            )
        except (KeyError, ValueError, FileNotFoundError) as exc:
            return jsonify({"error": f"Descrieri comerciale invalide: {exc}"}), 400
        if description_plan["missing"]:
            examples = ", ".join(
                f"{row.get('material')} ({'/'.join(row.get('missing_languages') or [])})"
                for row in description_plan["missing"][:5]
            )
            hs_session_event(
                {
                    "type": "log",
                    "level": "warn",
                    "msg": (
                        "Sar "
                        f"{len(description_plan['missing'])} materiale fara text "
                        "DE/EN in referinta"
                        + (f": {examples}" if examples else "")
                        + ". Continui clasificarea pentru restul."
                    ),
                }
            )

    # The variant carries the selection criteria, so it is applied per group.
    variant = (payload.get("variant") or "").strip()
    if not variant:
        try:
            variant = scheme_info(result["scheme"])["variant"]
        except KeyError:
            variant = ""

    creds = None
    cred_in = payload.get("credentials") or {}
    if cred_in.get("username") and cred_in.get("password"):
        creds = {"username": cred_in["username"], "password": cred_in["password"]}

    # Semi-automatic runs start from an uploaded worklist, so the browser may
    # not be open yet; the job is queued and runs once the session is ready.
    sess = get_hs_session()
    if sess.state["state"] in ("stopped", "error"):
        sess.start(
            {"sap_url": hs["sap_url"], "headless": bool(payload.get("headless"))},
            creds,
            hs_session_event,
        )

    job_id, q, cancel = _hs_new_job(
        kind="classify",
        hs=hs,
        items=[{"scheme": result["scheme"], "variant": variant}],
        allow_commit=allow_commit,
        auto_classify=True,
        source_file=result.get("source", ""),
    )

    def progress(ev):
        enriched = dict(ev)
        if enriched.get("type") not in {"log", "step", "fatal", "_end_"}:
            enriched.setdefault("index", 1)
            enriched.setdefault("scheme", result["scheme"])
        if enriched.get("type") == "job-end":
            enriched["remaining_worklist"] = _remaining_worklist_artifact(
                job_id,
                scheme=result["scheme"],
                index=1,
                source_path=result.get("_source_path", ""),
                classify_summary=enriched.get("summary"),
            )
        _hs_record_and_queue(job_id, q, enriched)

    sess.submit(
        "classify",
        {
            "groups": groups,
            # Selected HS groups come from the analysis "ready" set, so
            # replaying the full blocked list here only floods live events.
            "blocked": [],
            "scheme": result["scheme"],
            "variant": variant,
            "allow_commit": allow_commit,
            "description_plan": description_plan,
            "description_languages": hs["description_languages"],
            "description_url": commercial_description_url(hs["sap_url"]),
            "classification_url": hs["sap_url"],
            "description_display_maintained": hs["description_display_maintained"],
            "display_all_products": hs.get("display_all_products") is True,
        },
        progress,
        cancel,
    )
    return jsonify(
        {
            "job": job_id,
            "groups": len(groups),
            "products": sum(g["count"] for g in groups),
            "variant": variant,
            "allow_commit": allow_commit,
            "dry_run": not allow_commit,
            "descriptions_required": allow_commit,
            "description_products": len(description_plan["items"])
            if description_plan
            else 0,
            "report_url": f"/apps/hs/reports?report={job_id}",
        }
    )


@app.post("/hs/descriptions")
def hs_descriptions():
    """Maintain DE/EN for selected worklist groups without rewriting HS codes."""
    payload = request.get_json(force=True) or {}
    analysis_id = str(payload.get("analysis") or "")
    result = HS_ANALYSES.get(analysis_id)
    if not result:
        return jsonify({"error": "unknown analysis"}), 404

    hs_codes = payload.get("hs_codes")
    if not isinstance(hs_codes, list):
        return jsonify(
            {"error": "Trimite explicit lista grupurilor TECDOC/HS selectate."}
        ), 400
    groups = selected_groups(result, hs_codes)
    if not groups:
        return jsonify({"error": "Niciun grup selectat pentru descrieri."}), 400

    hs = hs_config()
    requested_commit = payload.get("allow_commit") is True
    allow_commit = requested_commit and hs.get("allow_commit") is True
    if not allow_commit:
        return jsonify(
            {
                "error": "Mentinerea manuala a descrierilor scrie in SAP. "
                "Bifeaza salvarea si verifica permisiunea locala."
            }
        ), 400

    try:
        reference = load_reference(hs)
        description_plan = build_description_plan(
            groups, reference, hs["description_languages"]
        )
    except (KeyError, ValueError, FileNotFoundError) as exc:
        return jsonify({"error": f"Descrieri comerciale invalide: {exc}"}), 400
    if not description_plan["items"]:
        return jsonify(
            {
                "error": "Nicio descriere DE/EN pregatita pentru grupurile selectate.",
                "missing": len(description_plan.get("missing") or []),
                "skipped": len(description_plan.get("skipped") or []),
            }
        ), 400
    if description_plan["missing"]:
        examples = ", ".join(
            f"{row.get('material')} ({'/'.join(row.get('missing_languages') or [])})"
            for row in description_plan["missing"][:5]
        )
        hs_session_event(
            {
                "type": "log",
                "level": "warn",
                "msg": (
                    "Sar "
                    f"{len(description_plan['missing'])} materiale fara text "
                    "DE/EN in referinta"
                    + (f": {examples}" if examples else "")
                    + ". Continui descrierile pentru restul."
                ),
            }
        )

    creds = None
    cred_in = payload.get("credentials") or {}
    if cred_in.get("username") and cred_in.get("password"):
        creds = {"username": cred_in["username"], "password": cred_in["password"]}

    sess = get_hs_session()
    if sess.state["state"] in ("stopped", "error"):
        sess.start(
            {"sap_url": hs["sap_url"], "headless": bool(payload.get("headless"))},
            creds,
            hs_session_event,
        )

    job_id, q, cancel = _hs_new_job(
        kind="descriptions",
        hs=hs,
        items=[{"scheme": result["scheme"], "variant": result.get("variant") or ""}],
        allow_commit=True,
        auto_classify=False,
        source_file=result.get("source", ""),
    )

    def progress(ev):
        enriched = dict(ev)
        if enriched.get("type") not in {"log", "step", "fatal", "_end_"}:
            enriched.setdefault("index", 1)
            enriched.setdefault("scheme", result["scheme"])
        _hs_record_and_queue(job_id, q, enriched)

    sess.submit(
        "descriptions",
        {
            "scheme": result["scheme"],
            "description_plan": description_plan,
            "description_languages": hs["description_languages"],
            "description_url": commercial_description_url(hs["sap_url"]),
            "classification_url": hs["sap_url"],
            "description_display_maintained": hs["description_display_maintained"],
        },
        progress,
        cancel,
    )
    return jsonify(
        {
            "job": job_id,
            "groups": len(groups),
            "products": sum(g["count"] for g in groups),
            "description_products": len(description_plan["items"]),
            "allow_commit": True,
            "dry_run": False,
            "descriptions_only": True,
            "report_url": f"/apps/hs/reports?report={job_id}",
        }
    )


@app.post("/hs/queue/start")
def hs_queue_start():
    """Process several (scheme, variant) pairs one after another.

    For each queue item: apply the variant, set the numbering scheme, execute
    and export the worklist, analyze it against the reference workbook, and
    optionally maintain every ready group before moving to the next item.
    """
    payload = request.get_json(force=True) or {}
    raw_items = payload.get("items") or []
    if not raw_items:
        return jsonify({"error": "Adauga cel putin o schema in coada."}), 400

    items = []
    for raw in raw_items:
        scheme = (raw.get("scheme") or "").strip()
        try:
            info = scheme_info(scheme)
        except KeyError as exc:
            return jsonify({"error": str(exc)}), 400
        variant = (raw.get("variant") or info["variant"]).strip()
        items.append({"scheme": info["code"], "variant": variant})

    cfg = load_config()
    hs = hs_config(cfg)
    ref_path = Path(hs["reference_file"])
    if not ref_path.is_absolute():
        ref_path = ROOT / ref_path
    if not ref_path.exists():
        return jsonify({"error": f"Fisierul de referinta lipseste: {ref_path}"}), 400
    try:
        reference = load_reference(hs)
        reference.validate_description_columns(hs["description_languages"])
    except Exception as exc:
        return jsonify({"error": f"Fisier de referinta invalid: {exc}"}), 400

    creds = None
    cred_in = payload.get("credentials") or {}
    if cred_in.get("username") and cred_in.get("password"):
        creds = {"username": cred_in["username"], "password": cred_in["password"]}

    auto_classify = payload.get("auto_classify") is True
    requested_commit = payload.get("allow_commit") is True
    allow_commit = auto_classify and requested_commit and hs.get("allow_commit") is True

    sess = get_hs_session()
    if sess.state["state"] in ("stopped", "error"):
        sess.start(
            {"sap_url": hs["sap_url"], "headless": bool(payload.get("headless"))},
            creds,
            hs_session_event,
        )

    job_id, q, cancel = _hs_new_job(
        kind="queue",
        hs=hs,
        items=items,
        allow_commit=allow_commit,
        auto_classify=auto_classify,
    )

    def progress(ev):
        if ev.get("type") == "scheme-analysis":
            analysis_id = store_analysis(
                ev["result"],
                ev.get("file", ""),
                ev.get("path"),
            )
            analysis_event = {
                "type": "scheme-analysis",
                "index": ev.get("index"),
                "scheme": ev.get("scheme"),
                **analysis_response(analysis_id, ev["result"]),
            }
            _hs_record_and_queue(job_id, q, analysis_event)
            return
        if ev.get("type") == "scheme-end":
            enriched = dict(ev)
            source_path = enriched.pop("source_path", "")
            enriched["remaining_worklist"] = _remaining_worklist_artifact(
                job_id,
                scheme=str(enriched.get("scheme") or "scheme"),
                index=int(enriched.get("index") or 1),
                source_path=source_path,
                classify_summary=enriched.get("classify_summary"),
            )
            _hs_record_and_queue(job_id, q, enriched)
            return
        _hs_record_and_queue(job_id, q, ev)

    sess.submit(
        "queue",
        {
            "items": items,
            "reference_path": str(ref_path),
            "reference_sheet": hs.get("reference_sheet"),
            "auto_classify": auto_classify,
            "allow_commit": allow_commit,
            "description_languages": hs["description_languages"],
            "description_url": commercial_description_url(hs["sap_url"]),
            "classification_url": hs["sap_url"],
            "description_display_maintained": hs["description_display_maintained"],
            "display_all_products": hs.get("display_all_products") is True,
        },
        progress,
        cancel,
    )
    return jsonify(
        {
            "job": job_id,
            "items": items,
            "auto_classify": auto_classify,
            "allow_commit": allow_commit,
            "dry_run": not allow_commit,
            "descriptions_required": allow_commit,
            "report_url": f"/apps/hs/reports?report={job_id}",
        }
    )


@app.post("/hs/stop/<job>")
def hs_stop(job):
    evt = HS_JOB_CANCEL.get(job)
    if not evt:
        return jsonify({"error": "unknown job"}), 404
    evt.set()
    return jsonify({"status": "cancel requested"})


@app.get("/hs/events/<job>")
def hs_events(job):
    q = HS_JOBS.get(job)
    if not q:
        return jsonify({"error": "unknown job"}), 404

    def stream():
        while True:
            ev = q.get()
            if ev.get("type") == "_end_":
                yield "event: end\ndata: {}\n\n"
                break
            yield f"data: {json.dumps(ev, ensure_ascii=False)}\n\n"

    return Response(stream(), mimetype="text/event-stream")


_SHUTDOWN_LOCK = threading.Lock()
_SHUTDOWN_SCHEDULED = False


def _is_local_request() -> bool:
    addr = (request.remote_addr or "").strip().lower()
    return addr in {"127.0.0.1", "::1", "localhost"}


def _schedule_process_exit(delay_s: float = 0.4) -> None:
    def _exit() -> None:
        time.sleep(delay_s)
        os._exit(0)

    threading.Thread(target=_exit, name="app-shutdown", daemon=True).start()


@app.post("/shutdown")
def shutdown_app():
    """Stop browsers and the local Python process so run.bat can start again."""
    if not _is_local_request():
        return (
            jsonify(
                {
                    "ok": False,
                    "error": "Oprirea aplicatiei este permisa doar de pe acest calculator.",
                }
            ),
            403,
        )

    global _SHUTDOWN_SCHEDULED
    with _SHUTDOWN_LOCK:
        already = _SHUTDOWN_SCHEDULED
        _SHUTDOWN_SCHEDULED = True

    hs_cleanup: dict = {"killed_pids": [], "removed": []}
    desc_cleanup: dict = {"killed_pids": [], "removed": []}
    if not already:
        try:
            get_hs_session().stop()
        except (RuntimeError, OSError):
            pass
        try:
            get_session().stop()
        except (RuntimeError, OSError):
            pass
        hs_cleanup = cleanup_hs_lock(kill_chrome=True)
        desc_cleanup = cleanup_user_data_lock(kill_chrome=True)
        _schedule_process_exit()

    return jsonify(
        {
            "ok": True,
            "shutting_down": True,
            "hs": hs_cleanup,
            "descriptions": desc_cleanup,
        }
    )


if __name__ == "__main__":
    from waitress import serve

    print("Starting server on http://localhost:5000 (waitress) ...")
    # threads must be high enough to cover SSE streams (each open stream
    # holds one worker thread) plus regular requests.
    serve(app, host="127.0.0.1", port=5000, threads=16, ident="FT6AUTO")
