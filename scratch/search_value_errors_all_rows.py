import os
import sys
import io
import time
import gspread
from google.oauth2.service_account import Credentials

# Fix encoding for Windows
os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except AttributeError:
    pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PARENT_DIR = os.path.dirname(BASE_DIR)
JSON_FILE = os.path.join(PARENT_DIR, 'credentials.json')

keys = {
    '1sUboaLTIeNTsbG56Re70-Xt5M7pHd8r4Neh-k4vzFt4': '2026 NTB - BÁO CÁO VẬN HÀNH',
    '1DAwY-46twFrHIs77R4p4IMuIZ6JTE-e58Aj-9Kcr5Jk': 'NTB - BÁO CÁO VẬN HÀNH'
}

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def main():
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    
    for key, name in keys.items():
        print(f"\n=========================================")
        print(f"Searching ALL rows in spreadsheet: {name} ({key})")
        try:
            sh = gc_client.open_by_key(key)
            for ws in sh.worksheets():
                time.sleep(2)  # Avoid rate limits
                try:
                    # Fetch all values
                    data = ws.get_all_values()
                    errs = []
                    for r_idx, row in enumerate(data):
                        for c_idx, val in enumerate(row):
                            if '#VALUE!' in str(val):
                                errs.append((r_idx+1, c_idx+1, val))
                    if errs:
                        print(f"  ❌ Sheet '{ws.title}' has {len(errs)} #VALUE! errors (first 10 shown):")
                        for r, c, val in errs[:10]:
                            col_letter = gspread.utils.rowcol_to_a1(r, c)[:-1]
                            row_preview = data[r-1][:5] if r-1 < len(data) else []
                            print(f"    Row {r} Col {col_letter} ({c}): {val} | Row preview: {row_preview}")
                    else:
                        print(f"  ✅ Sheet '{ws.title}' is CLEAN (0 #VALUE! errors, checked {len(data)} rows)")
                except Exception as ex_ws:
                    print(f"  Error reading '{ws.title}': {ex_ws}")
        except Exception as e:
            print(f"Error opening sheet: {e}")

if __name__ == "__main__":
    main()
