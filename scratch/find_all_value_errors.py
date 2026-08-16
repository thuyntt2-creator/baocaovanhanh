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
    '1DAwY-46twFrHIs77R4p4IMuIZ6JTE-e58Aj-9Kcr5Jk': 'NTB - BÁO CÁO VẬN HÀNH',
    '1WCzgao34cA_SttyB9ytHfE1qKTNl_3iFqDbEfw3lbyU': 'Aging >5 ngày',
    '1l2jZGLFoqxta2jz1RRJDTa1x15nHmXxGa2ZXhWXjM4M': 'Aging >5 ngày - follow gán',
    '1PjzFqJO-wkQ8SNsPHD721_CbPr6c_ArZKuGGU6KqDZg': 'NTB- FOLLOW OFF Tuyến',
    '1JZ1eRerRqrpwjZ4HBevQunjd8VquM_cvPFz12TaJfMQ': 'Dash Board'
}

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def main():
    print("Connecting to Google Sheets...")
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    
    for key, name in keys.items():
        print(f"\nScanning spreadsheet: {name} ({key})")
        try:
            sh = gc_client.open_by_key(key)
            worksheets = sh.worksheets()
            for ws in worksheets:
                # To speed up, we can fetch all values
                try:
                    data = ws.get_all_values()
                    errors = 0
                    sample = []
                    for r_idx, row in enumerate(data):
                        for c_idx, val in enumerate(row):
                            if '#VALUE!' in val:
                                errors += 1
                                if errors <= 3:
                                    col_letter = gspread.utils.rowcol_to_a1(r_idx+1, c_idx+1)[:-1]
                                    sample.append(f"Row {r_idx+1} Col {col_letter} ({c_idx+1}): {row[:5]}")
                    if errors > 0:
                        print(f"  ❌ Sheet '{ws.title}' has {errors} #VALUE! errors!")
                        for s in sample:
                            print(f"    - {s}")
                except Exception as ex_ws:
                    print(f"  Warning: failed to read sheet '{ws.title}': {ex_ws}")
        except Exception as e:
            print(f"Error opening sheet: {e}")

if __name__ == "__main__":
    main()
