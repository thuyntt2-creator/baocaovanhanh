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
SHEET_KEY = '1DAwY-46twFrHIs77R4p4IMuIZ6JTE-e58Aj-9Kcr5Jk'

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def main():
    print("Connecting to Google Sheets...")
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    sh = gc_client.open_by_key(SHEET_KEY)
    
    ws = sh.worksheet('raw')
    print("\nReading Row 1 formulas from 'raw' worksheet in 1DAw...")
    row1 = ws.get_values('A1:Z1', value_render_option='FORMULA')
    if row1:
        for idx, val in enumerate(row1[0]):
            col_letter = gspread.utils.rowcol_to_a1(1, idx+1)[:-1]
            print(f"Col {col_letter} ({idx+1}): {val}")
            
    print("\nReading Row 2 formulas from 'raw' worksheet in 1DAw...")
    row2 = ws.get_values('A2:Z2', value_render_option='FORMULA')
    if row2:
        for idx, val in enumerate(row2[0]):
            col_letter = gspread.utils.rowcol_to_a1(2, idx+1)[:-1]
            print(f"Col {col_letter} ({idx+1}): {val}")

if __name__ == "__main__":
    main()
