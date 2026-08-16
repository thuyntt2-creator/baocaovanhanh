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
    
    # Print spreadsheet properties/metadata
    print(f"Spreadsheet title: {sh.title}")
    # gspread spreadsheet object has metadata
    metadata = sh.fetch_sheet_metadata()
    properties = metadata.get('properties', {})
    print(f"Spreadsheet Locale: {properties.get('locale')}")
    print(f"Spreadsheet Timezone: {properties.get('timeZone')}")
    
    ws = sh.worksheet('DB')
    print("\nReading cells from 'DB' worksheet...")
    
    # Read values
    rows_val = ws.get_values('A1:I5')
    print("Values (get_values):")
    for r_idx, row in enumerate(rows_val):
        print(f"  Row {r_idx+1}: {row}")
        
    # Read formulas
    rows_form = ws.get_values('A1:U5', value_render_option='FORMULA')
    print("\nFormulas (FORMULA render):")
    for r_idx, row in enumerate(rows_form):
        # print length and first few items
        print(f"  Row {r_idx+1} (len {len(row)}): {row[:12]}")
        if len(row) > 12:
            print(f"    Rest: {row[12:]}")

if __name__ == "__main__":
    main()
