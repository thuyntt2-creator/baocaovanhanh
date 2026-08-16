import os
import sys
import io
import pandas as pd

os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except AttributeError:
    pass

BASE_DIR = r"c:\Users\lap4all\Documents\Auto report"

def inspect_excel(filename):
    path = os.path.join(BASE_DIR, filename)
    if os.path.exists(path):
        try:
            xl = pd.ExcelFile(path)
            print(f"\n======================================")
            print(f"File: {filename}")
            print(f"Sheets: {xl.sheet_names}")
            for sheet in xl.sheet_names[:10]: # Read up to 10 sheets
                # read header of each sheet
                df = pd.read_excel(path, sheet_name=sheet, nrows=5)
                print(f"  Sheet '{sheet}' columns: {list(df.columns)}")
                print(f"  First row: {df.iloc[0].to_dict() if len(df) > 0 else 'Empty'}")
        except Exception as e:
            print(f"Error reading {filename}: {e}")
    else:
        print(f"File {filename} does not exist.")

def main():
    inspect_excel("manual_raw.xlsx")

if __name__ == "__main__":
    main()
