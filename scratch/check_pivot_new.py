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
    
    new_key = '1vCxSTNgSpO9ETvVRElGyuGc7lnx7LxLRhAB4-lJMHLU'
    print(f"--- Checking PIVOT tab of {new_key} ---")
    try:
        sh = gc_client.open_by_key(new_key)
        ws = sh.worksheet("PIVOT")
        rows = ws.get_all_values()
        print(f"Total rows in PIVOT: {len(rows)}")
        print("Rows 1 to 20:")
        for idx, r in enumerate(rows[:20]):
            print(f"Row {idx+1}: {r}")
    except Exception as e:
        print("Error:", e)

if __name__ == '__main__':
    main()
