import os
import sys
import json
import gspread
from google.oauth2.service_account import Credentials

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

SHEET_KEY = '1PjzFqJO-wkQ8SNsPHD721_CbPr6c_ArZKuGGU6KqDZg'
CRED_FILE = r'c:\Users\lap4all\Documents\Auto report\credentials.json'

scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
creds = Credentials.from_service_account_file(CRED_FILE, scopes=scopes)
gc = gspread.authorize(creds)
sh = gc.open_by_key(SHEET_KEY)

ws = sh.worksheet("Đang OFF")
rows = ws.get_all_values()

headers = rows[0]
data_rows = rows[1:]

am_data = {}
for idx, r in enumerate(data_rows):
    if not any(r): continue
    # AM is col K (index 10)
    am = r[10].strip() if len(r) > 10 else "Unknown"
    if am not in am_data:
        am_data[am] = []
    am_data[am].append(r)

print(f"Total AMs with OFF data: {len(am_data)}")
for am, items in am_data.items():
    print(f"\n=== AM: {am} ({len(items)} tuyen) ===")
    for item in items:
        province = item[0]
        district = item[1]
        ward = item[2]
        ward_id = item[3]
        post_office = item[4]
        result = item[6]
        cap_down = item[7]
        off_from = item[8]
        off_to = item[9]
        print(f"  - [{result}] {province} | {district} | {ward} (ID: {ward_id}) | Kho: {post_office} | CapDown: {cap_down} | Tắt: {off_from} -> Mở: {off_to}")
