import os
import sys
import io
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
        print(f"\nSearching in spreadsheet: {name} ({key})")
        try:
            sh = gc_client.open_by_key(key)
            for ws in sh.worksheets():
                try:
                    # Let's read A1:Z100 to be fast and safe
                    cells = ws.get_values('A1:U100')
                    for r_idx, row in enumerate(cells):
                        for c_idx, val in enumerate(row):
                            if '2026-07-09' in val:
                                print(f"  Found in ws '{ws.title}' Row {r_idx+1} Col {c_idx+1}: {row[:10]}")
                except Exception as ex_ws:
                    print(f"  Error reading '{ws.title}': {ex_ws}")
        except Exception as e:
            print(f"Error opening sheet: {e}")

if __name__ == "__main__":
    main()
