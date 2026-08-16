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
    
    target_ids = ['20804000', '23034000', '20779000', '22427000', '1084', '21639000']
    
    for tid in target_ids:
        print(f"\nSearching for '{tid}'...")
        found = False
        for ws in sh.worksheets():
            try:
                cells = ws.findall(tid)
                if cells:
                    found = True
                    print(f"  Found in sheet '{ws.title}': {len(cells)} times (first match at Row {cells[0].row}, Col {cells[0].col})")
            except Exception as e:
                print(f"  Error searching in '{ws.title}': {e}")
        if not found:
            print("  Not found in any worksheet.")

if __name__ == "__main__":
    main()
