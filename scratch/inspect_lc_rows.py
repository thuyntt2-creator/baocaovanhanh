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
ws = sh.worksheet('stuck')
vals = ws.get_all_values()

headers = vals[0]
print("Headers:", headers)
print("Total rows:", len(vals) - 1)

statuses = {}
ton_vals = {}
ams = {}

for r in vals[1:]:
    if len(r) >= 9:
        st = r[4].strip()
        ton = r[5].strip()
        am = r[8].strip()
        statuses[st] = statuses.get(st, 0) + 1
        ton_vals[ton] = ton_vals.get(ton, 0) + 1
        ams[am] = ams.get(am, 0) + 1

print("\nStatuses:", statuses)
print("\nTồn đọng values sample:", list(ton_vals.items())[:15])
print("\nAMs:", ams)
