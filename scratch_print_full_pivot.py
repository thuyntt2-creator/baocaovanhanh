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
    sh = gc_client.open_by_key('1l2jZGLFoqxta2jz1RRJDTa1x15nHmXxGa2ZXhWXjM4M')
    
    ws = sh.worksheet("PIVOT")
    rows = ws.get_all_values()
    for idx, r in enumerate(rows):
        print(f"Row {idx+1}: {r}")

if __name__ == '__main__':
    main()
