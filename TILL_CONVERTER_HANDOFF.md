# Till Converter Debugging Handoff

## The Core Problem

The Till Converter is returning **0 employees and 0 dates** because the Revel PDF's internal structure doesn't match what the extraction logic expects.

### Why: The PDF Isn't Really a Table

To a human reader, the Revel PDF looks like a clean structured table with columns: `POS Station | Till Owner | Opened | Held | Closed | Declared | Expected | Variance`.

**But to pdfplumber, it's messy text.** When pdfplumber extracts the PDF:

1. **Rows are broken across newlines** — cells that appear to be neat columns in the rendered PDF are actually separated by multiple `\n` characters
2. **Column boundaries don't exist** — the data is essentially one long stream of text, not discrete columns
3. **Whitespace is inconsistent** — names, IDs, dates, and numbers may have extra spaces, tabs, or line breaks between them
4. **Timestamps attached to dates** — dates come with times (e.g., `02/23/26 10:41 am`), which the original parsing didn't handle

### What the Original Code Expected

The old `extract_pdf_data()` function assumed:
- `page.extract_table()` would cleanly return rows as a list of lists, with 8 columns each
- Each cell would contain exactly one piece of data (one name, one date, one money value)
- Row indexing would work: `row[0]` = station, `row[1]` = till owner, etc.

**This assumption breaks because the PDF is not a structured table — it's rendered text.**

---

## Current Debugging Status

**File:** `tools/till_converter.py`, function `extract_pdf_data()` (lines ~298–360)

**What was added:**
- Changed from `page.extract_table()` (singular) to `page.extract_tables()` (plural) to handle multiple table sections per page
- Added detailed `[DEBUG]` print statements to the terminal that show:
  - How many pages the PDF has
  - How many tables are found on each page
  - How many rows per table
  - Why each row is being rejected (not enough money values, no date found, no name found, etc.)
  - When a row is successfully extracted (with employee name and date)

**How to see the debug output:**
- Run `python main.py` in a terminal (not the `.vbs`)
- Upload the Revel PDF
- Watch the terminal for `[DEBUG]` lines — this will show exactly where the parsing is failing

---

## What Needs to Happen Next

### Step 1: Get Real Debug Data
Run the tool with the actual Revel PDF and share the **full terminal output** (the `[DEBUG]` lines). This will tell us:
- Is `extract_tables()` finding any tables?
- How many rows are in each table?
- At what step are valid rows being rejected (money check, date check, name check)?

### Step 2: Switch to Line-by-Line Parsing (if needed)

If `extract_tables()` still returns 0 rows, we'll need to abandon the table approach and parse the PDF as raw text instead:

```python
# Pseudo-code for the fallback approach:
def extract_pdf_data_fallback(pdf_file):
    till_data = []
    with pdfplumber.open(BytesIO(pdf_file.read())) as pdf:
        for page in pdf.pages:
            # Get raw text instead of trying to extract a table
            text = page.extract_text()
            
            # Split into lines and scan each one for data patterns
            for line in text.split('\n'):
                # Use regex to hunt for:
                # - A date: \d{1,2}/\d{1,2}/\d{2,4}
                # - A name: ([A-Za-z]+) \((\d+)\) ([A-Za-z]+)
                # - Three money values: \$?\d+\.\d{2}
                
                # If all three patterns match, extract and append to till_data
```

This is a **more robust fallback** if the table structure is too messy.

### Step 3: Test Against Real Data

Once parsing works, verify:
- Employee names are extracted as "FirstName LastName" (not just "FirstName" or with ID number)
- Dates are normalized to `YYYY-MM-DD` format
- The $200 float adjustment is applied correctly
- The final output matches the expected shape (rows per date, columns per employee)

---

## Key Files and Functions

| File | Function | Purpose |
|---|---|---|
| `tools/till_converter.py` | `extract_pdf_data()` | Extracts till rows from PDF — **THIS IS WHERE THE PROBLEM IS** |
| `tools/till_converter.py` | `extract_template_dates()` | Reads dates from Excel template (this part works) |
| `tools/till_converter.py` | `process_till_data()` | Groups by date/employee and builds output table (this part works) |
| `tools/till_converter.py` | `_parse_name()` | Cleans "First (ID) Last" into "First Last" |
| `tools/till_converter.py` | `_parse_date()` | Converts MM/DD/YY to YYYY-MM-DD |

---

## Quick Reference: Regex Patterns Used

```python
# Date: MM/DD/YY or MM/DD/YYYY
r'(\d{1,2}/\d{1,2}/\d{2,4})'

# Name: First (ID) Last
r'([A-Za-z]+)\s*\(\d+\)\s*([A-Za-z]+)'

# Money: any decimal number (possibly negative)
r'-?\d+\.\d{2}'

# Header skip: look for these strings to reject non-data rows
"POS STATION" or "TILL OWNER"
```

---

## Expected Behavior (Once Fixed)

- Tool accepts a Revel PDF and HDOS Excel template
- User selects a store number
- Backend extracts all till records from the PDF (all employees, all dates)
- Dates are matched to the Excel template's date column
- Output is a table ready to copy-paste into Excel
- Status shows: "Conversion complete! 9 employees, 35 dates."

---

## Notes for the Next Session

- **Don't assume `extract_table()` works.** The PDF is not a table; it's rendered text.
- **Print the raw text output** (`page.extract_text()`) to see what pdfplumber actually sees.
- **The regex patterns are good** — they're not the bottleneck. The bottleneck is getting the raw text to scan.
- **Test against the actual Revel PDF** — generic test PDFs won't show the real issues.
