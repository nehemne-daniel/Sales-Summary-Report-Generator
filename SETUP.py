#!/usr/bin/env python3
"""
Store Summary Report Generator - SETUP SCRIPT
Run this ONCE to set up the program
"""

import os
import sys
import subprocess
import shutil
from pathlib import Path

print("\n" + "="*80)
print("STORE SUMMARY REPORT GENERATOR - SETUP")
print("="*80)

# Create data folder
script_dir = Path(__file__).parent
data_dir = script_dir / "data"
data_dir.mkdir(exist_ok=True)

print("\nStep 1: Installing Python packages...")
packages = ['flask', 'openpyxl', 'python-docx']

for package in packages:
    try:
        __import__(package if package != 'python-docx' else 'docx')
        print(f"  ✓ {package} already installed")
    except ImportError:
        print(f"  Installing {package}...")
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', '-q', package])
        print(f"  ✓ {package} installed")

print("\nStep 2: Setting up template...")

# Check if template exists
template_file = script_dir / "Store_Summary_Report_Template.docx"

if template_file.exists():
    # Copy template to data folder
    data_template = data_dir / "template.docx"
    shutil.copy2(template_file, data_template)
    print(f"  ✓ Template prepared")
    print(f"  Note: You can now delete 'Store_Summary_Report_Template.docx' from this folder")
else:
    print(f"  ⚠ Template file not found!")
    print(f"  Please place 'Store_Summary_Report_Template.docx' in the same folder as this script")
    print(f"  Then run setup again")
    input("\nPress Enter to close...")
    sys.exit(1)

print("\n" + "="*80)
print("✓ SETUP COMPLETE!")
print("="*80)
print("\nYou can now run 'Store_Report_Generator.py'")
print("You won't need to run this setup script again")
print("\n" + "="*80 + "\n")

input("Press Enter to close this window...")
