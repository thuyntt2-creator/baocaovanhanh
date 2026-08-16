import os
import sys
import io
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd

# Fix encoding for Windows
os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except AttributeError:
    pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
EXTERNAL_SHEET_KEY = '1rhBD3-QSxn2c3yYNnyuj-iaW26H6WIH4G5FKb5Iv394'

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def main():
    print(f"Connecting to external Google Sheet: {EXTERNAL_SHEET_KEY}...")
    try:
        creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
        gc_client = gspread.authorize(creds)
        sh = gc_client.open_by_key(EXTERNAL_SHEET_KEY)
        print(f"Successfully opened sheet: '{sh.title}'")
        
        # List worksheets
        for ws in sh.worksheets():
            print(f"- Worksheet: {ws.title} | Rows: {ws.row_count} | Cols: {ws.col_count}")
            
        # Try reading 'Cơ cấu vùng_new'
        ws_vung = sh.worksheet("Cơ cấu vùng_new")
        # Get dimensions or first 20 rows
        vals = ws_vung.get_values("A1:L20")
        print("\nFirst 10 rows in 'Cơ cấu vùng_new':")
        for i, row in enumerate(vals[:10]):
            print(f"  Row {i+1}: {row}")
            
        # Read the whole sheet
        all_vals = ws_vung.get_all_values()
        df = pd.DataFrame(all_vals[2:], columns=all_vals[1]) # Row 3 is start of data, Row 2 is headers?
        print(f"\nTotal rows in 'Cơ cấu vùng_new': {len(df)}")
        print("Columns:", df.columns.tolist())
    except Exception as e:
        print(f"Failed to access external sheet: {e}")

if __name__ == "__main__":
    main()
