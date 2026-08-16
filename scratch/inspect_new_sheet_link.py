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
print("Title:", sh.title)
for ws in sh.worksheets():
    print(f"Sheet: '{ws.title}', id: {ws.id}")
    vals = ws.get_all_values()
    if vals:
        print(f"  Row count: {len(vals)}")
        print(f"  Header: {vals[0][:10]}")
    else:
        print("  Empty sheet")
