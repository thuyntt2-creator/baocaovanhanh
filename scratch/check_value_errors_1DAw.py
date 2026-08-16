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
    print("Fetching all values from 'raw' in 1DAw...")
    data = ws.get_all_values()
    print(f"Total rows in 'raw': {len(data)}")
    
    # Let's find rows containing '#VALUE!'
    error_rows = []
    for idx, row in enumerate(data):
        if any('#VALUE!' in val for val in row):
            error_rows.append((idx + 1, row))
            
    print(f"\nFound {len(error_rows)} rows with #VALUE!")
    for row_num, row in error_rows[:15]:
        print(f"Row {row_num}: {row[:9]} ... Error cols:")
        for col_idx, val in enumerate(row):
            if '#VALUE!' in val:
                col_letter = gspread.utils.rowcol_to_a1(row_num, col_idx+1)[:-1]
                print(f"  Col {col_letter} ({col_idx+1}): {val}")

if __name__ == "__main__":
    main()
