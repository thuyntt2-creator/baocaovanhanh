import os
import io
import sys
import gspread
from google.oauth2.service_account import Credentials

# Fix encoding cho Windows
os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except AttributeError:
    pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
SHEET_KEY = '1PjzFqJO-wkQ8SNsPHD721_CbPr6c_ArZKuGGU6KqDZg'
scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

def inspect_new_sheet():
    print(f"📖 Connecting to sheet: {SHEET_KEY}...")
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    
    try:
        sh = gc_client.open_by_key(SHEET_KEY)
        print(f"Success! Title: '{sh.title}'")
        worksheets = sh.worksheets()
        print("Worksheets inside this sheet:")
        for ws in worksheets:
            print(f"  - '{ws.title}' (ID: {ws.id})")
            
        # Inspect first worksheet's headers and first few rows
        ws = worksheets[0]
        data = ws.get_all_values()
        print(f"\nWorksheet '{ws.title}' has {len(data)} rows.")
        if len(data) > 0:
            print("Headers:")
            print(data[0])
            print("Row 1 sample:")
            if len(data) > 1:
                print(data[1])
            print("Row 2 sample:")
            if len(data) > 2:
                print(data[2])
    except Exception as e:
        import traceback
        print(f"❌ Error inspecting sheet: {e}")
        traceback.print_exc()

if __name__ == "__main__":
    inspect_new_sheet()
