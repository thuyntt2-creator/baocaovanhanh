import sys
import io
import os
import gspread
from google.oauth2.service_account import Credentials

try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except Exception:
    pass

BASE_DIR = r"c:\Users\lap4all\Documents\Auto report"
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def main():
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    
    print("--- Checking 1WCzgao34cA_SttyB9ytHfE1qKTNl_3iFqDbEfw3lbyU ---")
    try:
        sh = gc_client.open_by_key('1WCzgao34cA_SttyB9ytHfE1qKTNl_3iFqDbEfw3lbyU')
        for ws in sh.worksheets():
            rows = ws.get_all_values()
            print(f"Sheet: '{ws.title}', Row count: {len(rows)}")
            if len(rows) > 0:
                print(f"   First row: {rows[0][:5]}")
            if len(rows) > 1:
                print(f"   Second row: {rows[1][:5]}")
    except Exception as e:
        print("Error:", e)

if __name__ == '__main__':
    main()
