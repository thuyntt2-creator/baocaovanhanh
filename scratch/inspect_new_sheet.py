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
    print(f"--- Reading sheet {new_key} ---")
    try:
        sh = gc_client.open_by_key(new_key)
        print("Title:", sh.title)
        print("Worksheets:")
        for ws in sh.worksheets():
            print(f" - {ws.title} (ID: {ws.id})")
    except Exception as e:
        print("Error reading new sheet:", e)

if __name__ == '__main__':
    main()
