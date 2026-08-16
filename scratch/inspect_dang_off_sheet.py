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
print(f"Total rows in 'Đang OFF': {len(rows)}")

for i, row in enumerate(rows):
    print(f"Row {i+1:02d}: {row}")
