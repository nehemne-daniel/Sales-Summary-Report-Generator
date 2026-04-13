# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Store Summary Report Generator — a Python Flask app that takes two Excel files (current week + last year same week) and a Word template to produce a formatted `.docx` store summary report. Designed for retail store managers.

## Commands

**One-time setup** (installs dependencies and copies template to `data/`):
```bash
python SETUP.py
```

**Run the app** (starts Flask at http://localhost:5000, opens browser automatically):
```bash
python Store_Report_Generator.py
```

Dependencies: `flask`, `openpyxl`, `python-docx` — installed by `SETUP.py` via pip. There is no `requirements.txt`.

## Architecture

Everything is in a single file: [Store_Report_Generator.py](Store_Report_Generator.py).

**`HTML_FORM`** (line 44) — the entire frontend as a minified HTML string, rendered via `render_template_string`. JavaScript handles file uploads, repair row management, date extraction from filenames, and POSTing to `/generate`.

**`/generate` POST endpoint** (line 53) — the core logic:
1. Receives two `.xlsx` uploads + form fields via `multipart/form-data`
2. Calls `extract_excel_data()` and `extract_hours_worked()` on both files
3. Loads `data/template.docx` with `python-docx`
4. Writes data directly into specific table cells and paragraphs by index
5. Returns the modified document as a file download

**Excel extraction** — uses fixed cell addresses and row ranges:
- `K3` → total sales
- Rows 90–110, column A/B → "Total Transactions" and "Average Check"
- Rows 93–110, column A/B → "Service Crew" row for hours worked

**Word template structure** — tables are accessed by index (zero-based):
- `doc.tables[0]` → store info (name, manager, weekending date)
- `doc.tables[1]` → sales data (current week rows 1 & 3, last year rows 3 & 4)
- `doc.tables[2]` → labor (hours worked, scheduled, difference)
- `doc.tables[5]` → 86 repairs (rows appended dynamically)
- Comments and Team Update sections are found by scanning `doc.paragraphs` for the text `'Comments:'` and `'Team Update'`

**Template file**: `data/template.docx` (created by `SETUP.py` from `Store_Summary_Report_Template.docx`). The template is never modified in place — it is loaded fresh on each `/generate` call and saved to a `BytesIO` buffer for download.
