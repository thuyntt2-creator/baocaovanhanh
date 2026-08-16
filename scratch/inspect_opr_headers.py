import os
import sys
import io
import gspread
from google.oauth2.service_account import Credentials

os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except AttributeError:
    pass

BASE_DIR = r"c:\Users\lap4all\Documents\Auto report"
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
SPREADSHEET_ID = "1B-QCbEnPpILFFEWPYheGdmkgYV9gSf4lAyQMlhzwOCM"

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def main():
    print("Connecting to Google Sheets...")
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    sh = gc_client.open_by_key(SPREADSHEET_ID)
    
    print(f"Opened spreadsheet: {sh.title}")
    
    # List worksheets
    print("Worksheets:")
    for ws in sh.worksheets():
        print(f" - {ws.title}")
        
    try:
        ws_opr = sh.worksheet("OPR")
        print(f"\nWorksheet 'OPR' row count: {ws_opr.row_count}")
        rows = ws_opr.get_values("A1:Z10")
        print("\nFirst 10 rows:")
        for idx, r in enumerate(rows):
            print(f"Row {idx+1}: {r}")
    except Exception as e:
        print(f"Error reading 'OPR': {e}")

if __name__ == "__main__":
    main()
