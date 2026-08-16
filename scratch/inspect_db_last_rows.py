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
    
    ws = sh.worksheet('DB')
    print("Reading first 10 rows of 'DB' worksheet...")
    rows_first = ws.get_values('A1:I10')
    for idx, row in enumerate(rows_first):
        print(f"Row {idx+1}: {row}")
        
    print("\nReading last 10 rows of 'DB' worksheet...")
    # Get total rows
    total_rows = ws.row_count
    print(f"Total rows in sheet DB: {total_rows}")
    # Read values from row total_rows-10 to total_rows
    # To be safe, let's fetch all values and look at the end
    all_vals = ws.get_all_values()
    print(f"Actual non-empty rows: {len(all_vals)}")
    for idx, row in enumerate(all_vals[-10:]):
        print(f"Row {len(all_vals) - 10 + idx + 1}: {row}")

if __name__ == "__main__":
    main()
