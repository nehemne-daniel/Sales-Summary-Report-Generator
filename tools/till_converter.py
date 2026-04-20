#!/usr/bin/env python3
"""
Till PDF to Excel Converter - Flask Blueprint
Mounted at /till-converter by main.py. Not meant to be run directly.

Extracts till data from a Revel Systems "Till List" PDF and formats it for
pasting into the HDOS Daily Employee and Store Variance Excel template.
"""

import re
from io import BytesIO
from datetime import datetime
from collections import defaultdict

from flask import Blueprint, render_template_string, request, jsonify
import pdfplumber


HTML_FORM = """<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Till PDF to Excel Converter</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #c62b27; background-image: url('https://fundraising.hotdogonastick.com/_next/static/media/section-bg.3d099ea1.jpg'); background-size: cover; background-attachment: scroll; min-height: 100vh; padding: 0; }
        .stripe { width: 100%; height: 40px; background: repeating-linear-gradient(90deg,#e53430 0px,#e53430 60px,#0063a9 60px,#0063a9 120px,#fed500 120px,#fed500 180px,#f1f1f1 180px,#f1f1f1 240px); }
        .container { max-width: 1000px; margin: 0 auto; background: white; border-radius: 12px 12px 0 0; box-shadow: 0 20px 60px rgba(0,0,0,0.3); }
        .header { background: #e53430; color: white; padding: 40px; text-align: center; }
        .header h1 { font-size: 2.5em; margin-bottom: 10px; }
        .header p { font-size: 1.1em; opacity: 0.9; }
        .content { padding: 40px; }
        .form-group { margin-bottom: 20px; }
        label { display: block; font-weight: 600; margin-bottom: 8px; color: #333; }
        input[type="text"], input[type="file"], select { width: 100%; padding: 12px; border: 2px solid #ddd; border-radius: 6px; font-size: 1em; font-family: inherit; }
        input:focus, select:focus { outline: none; border-color: #0063a9; box-shadow: 0 0 0 3px rgba(0, 99, 169, 0.1); }
        .two-column { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }
        .button-group { display: flex; gap: 10px; margin-top: 30px; }
        button { flex: 1; padding: 14px 28px; font-size: 1em; font-weight: 600; border: none; border-radius: 6px; cursor: pointer; transition: all 0.3s; }
        .btn-primary { background: #e53430; color: white; }
        .btn-primary:hover { transform: translateY(-2px); box-shadow: 0 10px 20px rgba(229, 52, 48, 0.3); }
        .btn-primary:disabled { opacity: 0.6; cursor: not-allowed; }
        .btn-secondary { background: #f0f0f0; color: #333; }
        .btn-secondary:hover { background: #e0e0e0; }
        .status { padding: 15px; border-radius: 6px; margin-bottom: 20px; display: none; }
        .status.success { background: #d4edda; color: #155724; border: 1px solid #c3e6cb; }
        .status.error { background: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; }
        .status.loading { background: #d1ecf1; color: #0c5460; border: 1px solid #bee5eb; }
        .output-section { margin-top: 40px; padding: 20px; background: #f9f9f9; border-radius: 6px; display: none; }
        .output-section.active { display: block; }
        .output-section h2 { color: #e53430; margin-bottom: 10px; }
        #tableOutput { overflow-x: auto; max-width: 100%; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; background: white; border: 1px solid #ddd; font-size: 0.9em; }
        table th, table td { padding: 8px 10px; text-align: right; border: 1px solid #ddd; }
        table th:first-child, table td:first-child { text-align: left; }
        table th { background: #e53430; color: white; font-weight: 600; }
        table tr:nth-child(even) { background: #f9f9f9; }
        table tr:hover { background: #f0f0f0; }
        .copy-button { margin-top: 20px; padding: 10px 20px; background: #0063a9; color: white; border: none; border-radius: 6px; cursor: pointer; }
        .copy-button:hover { background: #004e86; }
        .copied-indicator { margin-left: 10px; padding: 6px 12px; background: #28a745; color: white; border-radius: 4px; font-size: 0.9em; font-weight: 500; display: inline-block; }
        .info-text { margin-top: 10px; font-size: 0.9em; color: #666; font-style: italic; }
    </style>
</head>
<body>
    <div class="stripe"></div>
    <div class="container">
        <div class="header">
            <h1>Till PDF to Excel Converter</h1>
            <p>Convert Revel Systems till data to HDOS format</p>
        </div>
        <div class="content">
            <div id="status" class="status"></div>

            <form id="converterForm">
                <div class="form-group">
                    <label for="pdfFile">Revel Till List PDF *</label>
                    <input type="file" id="pdfFile" accept=".pdf" required>
                </div>

                <div class="button-group">
                    <button type="submit" class="btn-primary" id="submitBtn">Convert PDF to Data</button>
                    <button type="button" class="btn-secondary" onclick="location.reload();">Clear</button>
                </div>
            </form>

            <div id="outputSection" class="output-section">
                <h2>Copy-Paste Ready Data</h2>
                <p class="info-text">Select the table below and copy (Ctrl+C), then paste into your Excel template. You can add/remove employee columns as needed.</p>
                <div id="tableOutput"></div>
                <button class="copy-button" onclick="copyTableToClipboard()">Copy Table to Clipboard</button>
                <span id="copiedIndicator" class="copied-indicator" style="display: none;">✓ Copied</span>
            </div>
        </div>
    </div>
    <div class="stripe"></div>

    <script>
        function showStatus(message, type) {
            const status = document.getElementById('status');
            status.textContent = message;
            status.className = `status ${type}`;
            status.style.display = 'block';
        }

        document.getElementById('converterForm').addEventListener('submit', async (e) => {
            e.preventDefault();
            try {
                showStatus('Processing files...', 'loading');
                document.getElementById('submitBtn').disabled = true;

                const pdfFile = document.getElementById('pdfFile').files[0];

                if (!pdfFile) {
                    showStatus('Please select a PDF file', 'error');
                    document.getElementById('submitBtn').disabled = false;
                    return;
                }

                const formData = new FormData();
                formData.append('pdfFile', pdfFile);

                showStatus('Converting PDF...', 'loading');
                const response = await fetch('/till-converter/convert', { method: 'POST', body: formData });

                if (!response.ok) {
                    const error = await response.json();
                    showStatus(`Error: ${error.error}`, 'error');
                    document.getElementById('submitBtn').disabled = false;
                    return;
                }

                const data = await response.json();
                displayTable(data.table);
                document.getElementById('outputSection').classList.add('active');
                showStatus(`Conversion complete! ${data.employeeCount} employees, ${data.dateCount} dates.`, 'success');
                document.getElementById('submitBtn').disabled = false;
            } catch (error) {
                console.error('Error:', error);
                showStatus(`Error: ${error.message}`, 'error');
                document.getElementById('submitBtn').disabled = false;
            }
        });

        function displayTable(tableData) {
            if (!tableData || tableData.length === 0) {
                document.getElementById('tableOutput').innerHTML = '<p>No data to display.</p>';
                return;
            }

            const headerRow = tableData[0];
            const employees = [];
            const totalsColumns = ['Revel', 'Loomis', 'Variance'];

            Object.keys(headerRow).forEach(key => {
                if (key === 'Date' || key === '' || key === 'Total' || totalsColumns.includes(key)) {
                    return;
                }
                employees.push(key);
            });
            employees.sort();

            const orderedColumns = ['Date', ...employees, 'Total', '', 'Revel', 'Loomis', 'Variance'];

            let html = '<table><thead><tr>';
            orderedColumns.forEach(col => { html += `<th>${col}</th>`; });
            html += '</tr></thead><tbody>';
            tableData.slice(1).forEach(row => {
                html += '<tr>';
                orderedColumns.forEach(col => { html += `<td>${row[col] || ''}</td>`; });
                html += '</tr>';
            });
            html += '</tbody></table>';
            document.getElementById('tableOutput').innerHTML = html;
        }

        function copyTableToClipboard() {
            const table = document.querySelector('#tableOutput table');
            if (!table) { return; }
            let text = '';
            table.querySelectorAll('tr').forEach(row => {
                const cells = row.querySelectorAll('th, td');
                text += Array.from(cells).map(c => c.textContent).join('\\t') + '\\n';
            });
            navigator.clipboard.writeText(text).then(
                () => {
                    const indicator = document.getElementById('copiedIndicator');
                    indicator.style.display = 'inline-block';
                    setTimeout(() => { indicator.style.display = 'none'; }, 2000);
                },
                err => console.error('Failed to copy:', err)
            );
        }
    </script>
</body>
</html>
"""


bp = Blueprint('till_converter', __name__, url_prefix='/till-converter')


@bp.route('/')
def index():
    return render_template_string(HTML_FORM)


@bp.route('/convert', methods=['POST'])
def convert():
    try:
        pdf_file = request.files['pdfFile']
        pdf_data = extract_pdf_data(pdf_file)
        table_data = process_till_data(pdf_data)

        # Count unique employees and dates that actually have data.
        employees = {row for row in table_data[0].keys()} - {'Date', 'Total', 'Revel', 'Loomis', 'Variance', ''}
        date_count = max(0, len(table_data) - 2)  # minus header and totals rows
        return jsonify({
            'table': table_data,
            'employeeCount': len(employees),
            'dateCount': date_count,
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 400


# "Alex (01) Zamora" -> ("Alex", "Zamora")
_NAME_RE = re.compile(r'^\s*(.+?)\s*\((\d+)\)\s*(.+?)\s*$')


def _parse_name(till_owner):
    if not till_owner:
        return None
    m = _NAME_RE.match(till_owner)
    if m:
        first, _id, last = m.groups()
        return f"{first} {last}".strip()
    # Fall back to whatever's before the paren if the pattern doesn't match.
    before_paren = till_owner.split('(')[0].strip()
    return before_paren or None




def _parse_money(raw):
    """Parse currency-ish string to float, or None on failure."""
    if raw is None or raw == '':
        return None
    try:
        return float(str(raw).replace('$', '').replace(',', '').strip())
    except (ValueError, TypeError):
        return None


def extract_pdf_data(pdf_file):
    """Extract till rows from a Revel 'Till List' PDF by scanning raw text with regex.

    The PDF isn't a structured table—it's rendered text. We extract raw text and hunt
    for data patterns (dates, names, money values) rather than relying on table structure.
    """
    till_data = []
    pdf_file.seek(0)

    with pdfplumber.open(BytesIO(pdf_file.read())) as pdf:
        print(f"[DEBUG] PDF has {len(pdf.pages)} pages")

        for page_idx, page in enumerate(pdf.pages):
            text = page.extract_text()
            if not text:
                print(f"[DEBUG] Page {page_idx + 1}: no text extracted")
                continue

            lines = text.split('\n')
            print(f"[DEBUG] Page {page_idx + 1}: {len(lines)} lines of text")

            for line_idx, line in enumerate(lines):
                line = line.strip()
                if not line:
                    continue

                # Skip headers and metadata rows
                upper = line.upper()
                if any(x in upper for x in ["POS STATION", "TILL OWNER", "OPENED", "HELD", "CLOSED", "DATE", "DECLARED", "EXPECTED", "VARIANCE"]):
                    continue

                # Hunt for 3+ decimal numbers (Declared, Expected, Variance)
                money_matches = re.findall(r'-?\d+\.\d{2}', line)
                if len(money_matches) < 3:
                    continue

                # Hunt for a date in MM/DD/YY or MM/DD/YYYY
                date_match = re.search(r'(\d{1,2}/\d{1,2}/\d{2,4})', line)
                if not date_match:
                    continue

                # Hunt for "First (ID) Last" name pattern
                name_match = re.search(r'([A-Za-z\-]+)\s*\(\d+\)\s*([A-Za-z\-]+)', line)
                if not name_match:
                    continue

                # Parse the data
                try:
                    date_str_raw = date_match.group(1)
                    try:
                        date_str = datetime.strptime(date_str_raw, "%m/%d/%y").strftime("%Y-%m-%d")
                    except ValueError:
                        date_str = datetime.strptime(date_str_raw, "%m/%d/%Y").strftime("%Y-%m-%d")

                    # Revel includes the $200 float; subtract it from declared and expected
                    declared_val = float(money_matches[-3].replace(',', '')) - 200
                    expected_val = float(money_matches[-2].replace(',', '')) - 200
                    variance_val = float(money_matches[-1].replace(',', ''))

                    employee = f"{name_match.group(1)} {name_match.group(2)}".strip()

                    print(f"[DEBUG] Extracted: {employee} | {date_str} | ${declared_val:.2f} | ${expected_val:.2f} | ${variance_val:.2f}")

                    till_data.append({
                        'employee': employee,
                        'date': date_str,
                        'declared': declared_val,
                        'expected': expected_val,
                        'variance': variance_val,
                    })
                except Exception as e:
                    print(f"[DEBUG] Page {page_idx + 1}, Line {line_idx}: parse error: {e}")
                    continue

    print(f"[DEBUG] Total rows extracted: {len(till_data)}")
    return till_data




def _fmt_date(iso_date):
    """Convert YYYY-MM-DD to DD-Mmm format (e.g., 23-Feb, 1-Mar)."""
    d = datetime.strptime(iso_date, '%Y-%m-%d')
    return f"{d.day}-{d.strftime('%b')}"


def _fmt_money(val):
    if val == 0:
        return ""
    return f"${val:.2f}"


def process_till_data(till_data):
    """Group by date and employee using PDF dates as the source of truth.
    Return a list of row-dicts: dates on Y-axis, employee names on X-axis,
    with Revel/Loomis/Variance on the right (separated by an empty column).

    First element is the header row. Last element is a totals row.
    """
    # Variance summed per (date, employee).
    variance_by_de = defaultdict(lambda: defaultdict(float))
    # Declared / expected / variance summed per date.
    day_totals = defaultdict(lambda: {'declared': 0.0, 'expected': 0.0, 'variance': 0.0})

    employees = []
    seen = set()
    for entry in till_data:
        d, emp = entry['date'], entry['employee']
        variance_by_de[d][emp] += entry['variance']
        day_totals[d]['declared'] += entry['declared']
        day_totals[d]['expected'] += entry['expected']
        day_totals[d]['variance'] += entry['variance']
        if emp not in seen:
            employees.append(emp)
            seen.add(emp)

    # Sort dates chronologically, employees alphabetically
    sorted_dates = sorted(day_totals.keys())
    sorted_employees = sorted(employees)

    # Header row
    header = {'Date': 'Date'}
    for emp in sorted_employees:
        header[emp] = emp
    header['Total'] = 'Total'
    header[''] = ''  # Empty column separator
    header['Revel'] = 'Revel'
    header['Loomis'] = 'Loomis'
    header['Variance'] = 'Variance'

    rows = [header]
    emp_totals = defaultdict(float)
    grand = {'declared': 0.0, 'expected': 0.0, 'variance': 0.0}

    # Data rows, one per date
    for date in sorted_dates:
        row = {'Date': _fmt_date(date)}
        for emp in sorted_employees:
            v = variance_by_de[date].get(emp, 0.0)
            row[emp] = _fmt_money(v)
            emp_totals[emp] += v
        dt = day_totals[date]
        row['Total'] = f"${dt['variance']:.2f}"
        row[''] = ''
        row['Revel'] = f"${dt['declared']:.2f}"
        row['Loomis'] = f"${dt['expected']:.2f}"
        row['Variance'] = f"${dt['variance']:.2f}"
        grand['declared'] += dt['declared']
        grand['expected'] += dt['expected']
        grand['variance'] += dt['variance']
        rows.append(row)

    # Totals row
    totals = {'Date': 'TOTAL'}
    for emp in sorted_employees:
        totals[emp] = f"${emp_totals[emp]:.2f}"
    totals['Total'] = f"${grand['variance']:.2f}"
    totals[''] = ''
    totals['Revel'] = f"${grand['declared']:.2f}"
    totals['Loomis'] = f"${grand['expected']:.2f}"
    totals['Variance'] = f"${grand['variance']:.2f}"
    rows.append(totals)

    return rows
