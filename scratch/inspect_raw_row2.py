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
        print(f"\n=========================================")
        print(f"Spreadsheet: {name} ({key})")
        try:
            sh = gc_client.open_by_key(key)
            ws = sh.worksheet('raw')
            
            # Fetch A2:U2 values
            row2_val = ws.get_values('A2:U2')
            print(f"Row 2 values: {row2_val}")
            
            # Fetch A2:U2 formulas
            row2_form = ws.get_values('A2:U2', value_render_option='FORMULA')
            print(f"Row 2 formulas: {row2_form}")
            
            # Print details of Col E to U
            if row2_val:
                headers = ws.get_values('A1:U1')[0]
                for idx in range(4, len(row2_val[0])):
                    header = headers[idx] if idx < len(headers) else f"Col {idx+1}"
                    val = row2_val[0][idx]
                    form = row2_form[0][idx] if idx < len(row2_form[0]) else ""
                    print(f"  Col {idx+1} ({header}): Value='{val}', Formula='{form}'")
                    
        except Exception as e:
            print(f"Error: {e}")

if __name__ == "__main__":
    main()
