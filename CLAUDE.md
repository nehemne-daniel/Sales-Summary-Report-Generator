# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

HDOS Tools — a suite of Python Flask tools for HDOS retail store managers, served from a single web app with a main menu at `http://localhost:5000/`. Current tools:

- **Store Summary Report Generator** (`/store-report/`) — takes two Excel files (current week + last year) and a Word template, produces a formatted `.docx` weekly report.
- **Till PDF → Excel Converter** (`/till-converter/`) — converts a Revel Systems till-list PDF into copy-paste-ready data for the HDOS Daily Employee and Store Variance Excel template.

## Commands

**One-time setup** (installs dependencies and copies the Word template into `data/`):
```bash
python SETUP.py
```

**Run the app** (starts Flask at `http://localhost:5000`, opens the browser automatically to the menu):
```bash
python main.py
```

Dependencies: `flask`, `openpyxl`, `python-docx`, `pdfplumber` — installed by `SETUP.py` via pip. No `requirements.txt`.

## Architecture

Single Flask app on port 5000 (`main.py`); each tool is a **Flask Blueprint** under `tools/`. Adding a new tool = drop a new blueprint module in `tools/` and register it in `main.py`.

### Folder layout

```
main.py                              entry point — menu at /, registers blueprints
SETUP.py                             one-time package install + template copy
Run_Report.vbs                       double-clickable launcher for main.py
MENU_NOTES.md                        deferred UX decisions for the menu
templates/
  Store_Summary_Report_Template.docx original Word template (source for SETUP.py)
tools/
  __init__.py
  store_report_generator.py          Blueprint — store summary report
  till_converter.py                  Blueprint — till PDF converter
data/
  template.docx                      working copy created by SETUP.py
```

### Routing

| URL | Served by |
|---|---|
| `/` | [main.py](main.py) — menu page |
| `/store-report/` + `/store-report/generate` + `/store-report/debug` | [tools/store_report_generator.py](tools/store_report_generator.py) |
| `/till-converter/` + `/till-converter/convert` | [tools/till_converter.py](tools/till_converter.py) |

Each tool blueprint uses `render_template_string` with an inline `HTML_FORM` string (no Flask `templates/` folder is used — `templates/` in this project holds Word/Excel source templates, not Jinja templates). JS in each tool POSTs to its own blueprint-prefixed endpoint (e.g. `/store-report/generate`, not `/generate`).

### Store Report Generator internals

[tools/store_report_generator.py](tools/store_report_generator.py):

- **`/generate` POST** — receives two `.xlsx` uploads + form fields, calls `extract_excel_data()` and `extract_hours_worked()`, loads `data/template.docx`, writes into tables/paragraphs by index, returns the `.docx` as a download.
- **Excel extraction** uses fixed cell addresses and row ranges:
  - `K3` → total sales
  - Rows 90–110, column A/B → "Total Transactions" and "Average Check"
  - Rows 93–110, column A/B → "Service Crew" row for hours worked
- **Word template tables** (accessed by index):
  - `doc.tables[0]` → store info (name, manager, weekending date)
  - `doc.tables[1]` → sales data (current week rows 1 & 3, last year rows 3 & 4)
  - `doc.tables[2]` → labor (hours worked, scheduled, difference)
  - `doc.tables[5]` → 86 repairs (rows appended dynamically)
  - Comments and Team Update sections are found by scanning `doc.paragraphs` for the text `'Comments:'` and `'Team Update'`
- Template is loaded fresh each request and saved to a `BytesIO` — the on-disk template is never modified.
- `template_path` is resolved as `Path(__file__).parent.parent / "data" / "template.docx"`, i.e. relative to the project root, since the module now lives inside `tools/`.

### Till Converter internals

[tools/till_converter.py](tools/till_converter.py):

- **`/convert` POST** — receives a Revel `.pdf`, an HDOS template `.xlsx`, and a store number. Extracts rows with `pdfplumber`, aligns to the template's date list, returns a table of `{date, emp_variances..., revel_declared, loomis_expected, variance}` as JSON for the UI to render as a copy-pasteable HTML table.
- **PDF row schema** (Revel export): `POS Station, Till Owner, Opened, Held, Closed, Declared Cash, Expected Cash, Variance`.
- **Name parsing** uses `_NAME_RE` to convert `"Alex (01) Zamora"` → `"Alex Zamora"`.
- **Date parsing** tries `%m/%d/%y`, `%m/%d/%Y`, `%Y-%m-%d` and normalizes to ISO `YYYY-MM-DD`, so PDF dates (`02/23/26`) match template dates (which may be date objects or strings).
- **$200 adjustment** — Revel includes the starting till float in both declared and expected; the code subtracts `200` from each before reporting. This is a fixed system quirk, not dynamic.
- **Per-page headers** — `extract_pdf_data` skips any row whose first cell contains `"POS STATION"`, because Revel repeats the header row on every page of the table.
