import os
import io
import sys
import gspread
from google.oauth2.service_account import Credentials

# Fix encoding for Windows
os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except AttributeError:
    pass

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PARENT_DIR = os.path.dirname(SCRIPT_DIR)
JSON_FILE = os.path.join(PARENT_DIR, 'credentials.json')
scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

def inspect_all():
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    
    sheet_key = '1DAwY-46twFrHIs77R4p4IMuIZ6JTE-e58Aj-9Kcr5Jk'
    sh = gc_client.open_by_key(sheet_key)
    print(f"Spreadsheet: {sh.title}")
    
    for ws in sh.worksheets():
        print(f"\n--- Worksheet: {ws.title} (ID: {ws.id}) ---")
        try:
            # get first 2 rows
            rows = ws.get_values("A1:Z5")
            for idx, r in enumerate(rows):
                print(f"  Row {idx+1}: {r[:10]}")
        except Exception as e:
            print(f"  Error reading: {e}")

if __name__ == "__main__":
    inspect_all()
