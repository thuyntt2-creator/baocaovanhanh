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
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
SHEET_KEY = '1sUboaLTIeNTsbG56Re70-Xt5M7pHd8r4Neh-k4vzFt4'

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def main():
    print("Connecting to Google Sheets...")
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    sh = gc_client.open_by_key(SHEET_KEY)
    
    sheets_to_inspect = ['raw', 'DB', 'dbtts', 'rawtts']
    
    for sname in sheets_to_inspect:
        print(f"\n--- Worksheet: {sname} ---")
        try:
            ws = sh.worksheet(sname)
            # Fetch row 1 and row 2 with formulas
            vals = ws.get_values('A1:U3', value_render_option='FORMULA')
            for i, row in enumerate(vals):
                print(f"  Row {i+1}: {row[:12]}") # Print first 12 columns
        except Exception as e:
            print(f"  Error reading sheet '{sname}': {e}")

if __name__ == "__main__":
    main()
