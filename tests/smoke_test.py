"""Quick smoke test: run the SAP automation against 1 row from coodesctest.xlsx."""

import json
import sys
import threading
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "server"))

from excel_reader import read_materials  # type: ignore
from sap_automation import run_update  # type: ignore

cfg = json.loads((ROOT / "config.json").read_text(encoding="utf-8-sig"))
xlsx = ROOT / "coodesctest.xlsx"
if not xlsx.exists():
    raise SystemExit(
        f"[ERROR] {xlsx.name} lipseste. Acest script scrie REAL in SAP; "
        "pune un fisier de test langa run.bat inainte de a-l rula."
    )
excel_cfg = cfg.get("excel", {})
items = read_materials(
    str(xlsx),
    material_column=excel_cfg.get("material_column", "B"),
    languages=cfg.get("languages"),
    header_row=int(excel_cfg.get("header_row", 1)),
    sheet=excel_cfg.get("sheet"),
)
print(f"[INFO] read {len(items)} rows from {xlsx.name}")
items = items[:1]
if not items:
    raise SystemExit("[ERROR] no valid Excel rows for the active languages")
print("[INFO] running 1 item:", items[0])

cancel = threading.Event()


def cb(ev):
    print("  [evt]", ev)


res = run_update(items, cfg, progress_cb=cb, cancel_event=cancel, headless=False)
print("[RESULT]", json.dumps(res, indent=2, ensure_ascii=False))
