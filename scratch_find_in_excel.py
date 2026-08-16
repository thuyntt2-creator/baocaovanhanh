import os
import sys
import io
import pandas as pd
import openpyxl

# Fix encoding for Windows
os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except AttributeError:
    pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def main():
    target = 21624000
    target_str = '21624000'
    
    excel_files = [f for f in os.listdir(BASE_DIR) if f.endswith('.xlsx')]
    print("Excel files in workspace:", excel_files)
    
    for f in excel_files:
        path = os.path.join(BASE_DIR, f)
        print(f"\nScanning Excel file: {f} (size: {os.path.getsize(path)} bytes)...")
        try:
            # We will read each sheet in the excel file
            xl = pd.ExcelFile(path)
            for sheet_name in xl.sheet_names:
                print(f"  Reading sheet: {sheet_name}...")
                df = pd.read_excel(path, sheet_name=sheet_name)
                # Check if target is in any column
                for col in df.columns:
                    # check numeric or string match
                    matches_num = df[df[col] == target]
                    matches_str = df[df[col].astype(str).str.strip() == target_str]
                    if len(matches_num) > 0 or len(matches_str) > 0:
                        print(f"    Found '{target_str}' in column '{col}' in sheet '{sheet_name}'!")
                        print("    Sample rows:")
                        print(df[df[col] == target].head(2).to_string())
                        print(df[df[col].astype(str).str.strip() == target_str].head(2).to_string())
        except Exception as e:
            print(f"  Error reading {f}: {e}")

if __name__ == "__main__":
    main()
