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
    
    print("\nWorksheets present in the spreadsheet:")
    for ws in sh.worksheets():
        # Get dimensions
        print(f"- Title: {ws.title} | Rows: {ws.row_count} | Cols: {ws.col_count}")
        # Get first row headers
        try:
            row1 = ws.get_values("A1:Z1")
            if row1:
                print(f"  Headers: {row1[0]}")
            else:
                print("  Headers: Empty")
        except Exception as e:
            print(f"  Error reading headers: {e}")

if __name__ == "__main__":
    main()
