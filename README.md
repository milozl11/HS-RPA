# FT6AUTO — SAP GTS HS Classification & Customs Description Uploader

Automated batch update of material customs descriptions in SAP Fiori (GTS Foreign Trade).
Reads materials + multilingual descriptions from an Excel file, then drives SAP via browser automation.

The HS queue uses one reference workbook as its local TECDOC database. For a
committed scheme it now runs an atomic sequence:

1. download and analyze the SAP worklist;
2. preflight the tariff code plus both DE and EN descriptions for every ready product;
3. maintain and confirm every HS group in `Classify Products`;
4. maintain and confirm DE + EN in `Manage Customs Commercial Descriptions`;
5. start the next scheme only after both SAP stages are complete.

A missing SAP variant or empty worklist skips that scheme and continues the
queue. After a worklist analysis, the operator can start DE/EN maintenance
alone for the selected TECDOC/HS groups, without rewriting HS codes.

Any missing reference text, partial SAP report, unconfirmed save, or missing
product still stops the remaining committed queue after a failed description
stage.

Every HS run also creates a persistent report under `reports-hs/<job-id>/`.
For each scheme, the application saves a copy of the original SAP workbook
with only SAP-confirmed classified materials removed. Blocked, failed,
unselected, and dry-run materials remain in the workbook. The original sheet
layout, columns, preamble, styles, and any additional sheets are preserved.

## Requirements

- **Windows 10/11** (64-bit)
- **No Python installation required**
- **No internet required** (when using the pre-built bundle)

## Distribution Methods

### Method A: Pre-built Bundle (recommended for corporate environments)

No internet, no pip, no downloads needed. Works behind any firewall.

1. Get `FT6AUTO_v1.2.zip` from the person who built the bundle
2. **Extract** the ZIP to any local folder (e.g. `C:\Tools\FT6AUTO`)
3. **Double-click `run.bat`**
4. The web UI opens at **http://localhost:5000**

> To create the bundle: run `bundle.bat` on a machine where setup has already been completed.

### Method B: Clone + Auto-Setup (requires internet once)

For machines with internet access to pypi.org and python.org:

1. **Clone or download** this repository
2. **Double-click `run.bat`**
   - On first run, it automatically downloads Python + dependencies + Chromium
   - This takes 2–5 minutes and requires internet
3. Subsequent runs work fully offline

### Manual Setup (optional)

If `run.bat` auto-setup fails:

```
setup.bat
run.bat
```

## Project Structure

```
FT6AUTO/
├── run.bat              # START HERE - launches the app (auto-setup on first run)
├── setup.bat            # One-time setup: downloads Python + deps + Chromium
├── bundle.bat           # Creates the distributable ZIP
├── config.json          # SAP URL, systems, languages, column mappings
├── requirements.txt     # Python dependencies (used by setup.bat)
├── Referinta/           # TECDOC reference workbook - update this file manually
├── server/
│   ├── app.py              # Flask web server + REST API + SSE streaming
│   ├── sap_automation.py   # SAP browser automation engine (Playwright)
│   ├── hs_automation.py    # HS classification + description flows
│   ├── hs_worklist.py      # Worklist analysis and grouping
│   ├── hs_reference.py     # TECDOC reference lookup
│   ├── hs_reporting.py     # Persistent run reports and monthly stats
│   ├── excel_reader.py     # Excel file parser (openpyxl)
│   ├── static/             # Shared stylesheet and wizard script
│   └── templates/          # Launcher, HS, descriptions, reports, dashboard
├── tests/               # Offline test suites (developers only)
├── python-embed/        # (bundled) Portable Python 3.11.9
├── playwright-browsers/ # (bundled) Chromium for Playwright
├── uploads/             # (created at runtime) Uploaded Excel files
├── uploads-hs/          # (created at runtime) HS worklists
├── downloads/           # (created at runtime) Files exported from SAP
├── reports-hs/          # (created at runtime) Reports + remaining worklists
└── user-data-hs/        # (created at runtime) Browser profile data
```

The operator only ever needs **`run.bat`**. Everything else is either
configuration (`config.json`, `Referinta/`) or created automatically.

## Config

Edit `config.json` to change the SAP URL and column mappings. The integrated HS
flow requires `hs.description_languages` for both `DE` and `EN` (defaults: Excel
columns `J` and `K`). It derives the descriptions-app URL from the selected HS
system, so classification and descriptions stay on the same SAP host/client.

The reference is read automatically at server startup and cached by file path,
sheet, modification time, and size. Replacing/updating the configured workbook
causes it to be re-read on the next status check or job. The HS screen shows
whether the file was loaded, its TECDOC count, and the DE/EN mappings.

The HS screen is also the control center for the integrated workflow. It shows
the five guarded stages (reference, worklist, HS, descriptions, complete),
configuration health, and separate HS versus DE/EN results per scheme. **Setari
flux** manages the default scheme, optional reference sheet, the DE/EN column
and SAP-country mappings, and whether already maintained description records
are included. The original descriptions screen remains available as a separate
standalone application.

The application hub, integrated HS screen, and HS report page support Romanian
and English. The selection is shared between these pages and persists locally
in the browser.

The **Rapoarte / Reports** page separates factual stage outcomes from the raw
technical timeline. It shows per-scheme analysis, HS confirmations, DE/EN
confirmations, material-level results, the remaining SAP worklist, and a CSV
audit export. Report data survives a page reload or application restart.

## Notes

- `python-embed/`, `playwright-browsers/`, `uploads/`, `uploads-hs/`, `downloads/`, `reports-hs/`, and `user-data-hs/` are **not** in git (runtime state or too large for GitHub)
- Run from a **local drive** (C:\, D:\) for best performance — network drives are slow

## Handing the tool to a colleague

Use `bundle.bat`. Do **not** zip the working folder by hand: `user-data-hs/`
holds the live SAP browser profile and `reports-hs/`, `downloads/`, `uploads-hs/`
hold material data from previous runs. `bundle.bat` packs only the application,
the bundled runtime, and the reference workbook.

The recipient extracts the ZIP and double-clicks `run.bat`. Nothing needs to be
installed on their machine.

## Developer checks

```
python-embed\python.exe tests\run_all.py
```

Runs every offline suite (backend routes, HS data, descriptions, reporting, SAP
dialogs, template DOM references) without touching SAP.

`tests\smoke_test.py` is different: it performs a **real** SAP write and needs a
test workbook next to `run.bat`.
