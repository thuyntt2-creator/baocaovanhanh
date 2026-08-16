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

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def main():
    print("Connecting to Google Sheets...")
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    
    for key, name in [
        ('1sUboaLTIeNTsbG56Re70-Xt5M7pHd8r4Neh-k4vzFt4', '2026 NTB - BÁO CÁO VẬN HÀNH'),
        ('1DAwY-46twFrHIs77R4p4IMuIZ6JTE-e58Aj-9Kcr5Jk', 'NTB - BÁO CÁO VẬN HÀNH')
    ]:
        print(f"\n=========================================")
        print(f"Spreadsheet: {name} ({key})")
        try:
            sh = gc_client.open_by_key(key)
            ws = sh.worksheet("Ca1 - Ca2 - Tồn")
            print("Reading cells A1:U5...")
            vals = ws.get_values('A1:U5')
            for r_idx, row in enumerate(vals):
                print(f"  Row {r_idx+1}: {row}")
                
            print("\nReading formulas A1:U5...")
            forms = ws.get_values('A1:U5', value_render_option='FORMULA')
            for r_idx, row in enumerate(forms):
                print(f"  Row {r_idx+1}: {row}")
        except Exception as e:
            print(f"Error: {e}")

if __name__ == "__main__":
    main()
