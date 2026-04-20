#!/usr/bin/env python3
"""
HDOS Tools - main entry point.
Serves a menu at / and mounts each tool as a Flask blueprint.
"""

import sys
sys.dont_write_bytecode = True

import webbrowser
from threading import Thread
from pathlib import Path

# Fail fast if SETUP.py hasn't been run (needed for Store Report template).
project_root = Path(__file__).parent
template_path = project_root / "data" / "template.docx"

if not template_path.exists():
    print("\n" + "="*80)
    print("ERROR: Setup not complete!")
    print("="*80)
    print("\nPlease run SETUP.py first")
    print("\nSteps:")
    print("  1. Make sure 'Store_Summary_Report_Template.docx' is in the 'templates' folder")
    print("  2. Double-click 'SETUP.py'")
    print("  3. Then run this program")
    print("\n" + "="*80 + "\n")
    input("Press Enter to close...")
    sys.exit(1)

try:
    from flask import Flask, render_template_string
except ImportError as e:
    print(f"\nERROR: Missing required package: {e}")
    print("Please run SETUP.py first")
    input("Press Enter to close...")
    sys.exit(1)

from tools.store_report_generator import bp as store_report_bp
from tools.till_converter import bp as till_converter_bp

MENU_HTML = """<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>HDOS Tools</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #c62b27; background-image: url('https://fundraising.hotdogonastick.com/_next/static/media/section-bg.3d099ea1.jpg'); background-size: cover; background-attachment: scroll; min-height: 100vh; padding: 0; }
.stripe { width: 100%; height: 40px; background: repeating-linear-gradient(90deg,#e53430 0px,#e53430 60px,#0063a9 60px,#0063a9 120px,#fed500 120px,#fed500 180px,#f1f1f1 180px,#f1f1f1 240px); }
.container { max-width: 1000px; margin: 0 auto; background: white; border-radius: 12px 12px 0 0; box-shadow: 0 20px 60px rgba(0,0,0,0.3); }
.header { background: #e53430; color: white; padding: 40px; text-align: center; }
.header h1 { font-size: 2.5em; margin-bottom: 10px; }
.header p { font-size: 1.1em; opacity: 0.9; }
.content { padding: 40px; }
.tools-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }
@media (max-width: 700px) { .tools-grid { grid-template-columns: 1fr; } }
.tool-card { display: block; padding: 30px; background: #f9f9f9; border: 2px solid #ddd; border-radius: 8px; text-decoration: none; color: #333; transition: all 0.2s; }
.tool-card:hover { border-color: #0063a9; transform: translateY(-2px); box-shadow: 0 10px 20px rgba(0, 99, 169, 0.15); background: white; }
.tool-card h2 { color: #e53430; font-size: 1.4em; margin-bottom: 10px; }
.tool-card p { color: #555; font-size: 0.95em; line-height: 1.5; }
.footer-note { margin-top: 30px; text-align: center; color: #888; font-size: 0.9em; font-style: italic; }
</style>
</head>
<body>
<div class="stripe"></div>
<div class="container">
  <div class="header">
    <h1>HDOS Tools</h1>
    <p>Choose a tool to get started</p>
  </div>
  <div class="content">
    <div class="tools-grid">
      <a class="tool-card" href="/store-report/" target="_blank">
        <h2>Store Summary Report Generator</h2>
        <p>Generate a weekly store summary report from two Excel files (current week + last year) using the Word template.</p>
      </a>
      <a class="tool-card" href="/till-converter/" target="_blank">
        <h2>Till PDF &rarr; Excel Converter</h2>
        <p>Convert a Revel Systems till-list PDF into copy-paste-ready data for the HDOS Daily Employee &amp; Store Variance template.</p>
      </a>
    </div>
    <div class="footer-note">Each tool opens in a new tab &mdash; this menu will stay open.</div>
  </div>
</div>
<div class="stripe"></div>
</body>
</html>
"""

app = Flask(__name__)
app.register_blueprint(store_report_bp)
app.register_blueprint(till_converter_bp)


@app.route('/')
def menu():
    return render_template_string(MENU_HTML)


if __name__ == '__main__':
    print("\n" + "="*80)
    print("HDOS TOOLS")
    print("="*80)
    print("\nStarting web server...")
    print("Browser will open automatically...\n")

    def open_browser():
        import time
        time.sleep(1)
        webbrowser.open('http://localhost:5000')

    Thread(target=open_browser, daemon=True).start()

    try:
        app.run(debug=False, port=5000, use_reloader=False)
    except KeyboardInterrupt:
        print("\n\nShutting down...")
    except Exception as e:
        print(f"\nERROR: {e}")
        input("\nPress Enter to close...")
