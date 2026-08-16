import os
import io
import sys
import gspread
from google.oauth2.service_account import Credentials

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
SHEET_KEY = '1MjLW8NbD5ZjoOdd90myGv0i1NGAtlvScxebfAXMM1j8'

scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
gc = gspread.authorize(creds)

sh = gc.open_by_key(SHEET_KEY)

for ws_name in ['PIVOT', 'Bảng theo dõi nhóm']:
    try:
        ws = sh.worksheet(ws_name)
        print(f"\n--- {ws_name} ---")
        rows = ws.get_all_values()
        for r in rows[:20]:
            print(r)
    except Exception as e:
        print(f"Error {ws_name}: {e}")
