import openpyxl
import os
import glob
import sys

sys.stdout.reconfigure(encoding='utf-8')

downloads_dir = r"C:\Users\lap4all\Downloads"
xlsx_files = glob.glob(os.path.join(downloads_dir, "*.xlsx"))

print("=== Checking Excel sheets in Downloads ===")
for f in xlsx_files:
    if "~$" in f:
        continue
    try:
        wb = openpyxl.load_workbook(f, read_only=True)
        print(f"\nFile: {os.path.basename(f)}")
        print(f"  Sheets: {wb.sheetnames}")
        wb.close()
    except Exception as e:
        print(f"  Error reading {os.path.basename(f)}: {e}")
