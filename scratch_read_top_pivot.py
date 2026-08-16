import json
import sys
import io
import gspread
from google.oauth2.service_account import Credentials

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

JSON_FILE = r"c:\Users\lap4all\Documents\Auto report\credentials.json"
scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
gc_client = gspread.authorize(creds)

sh = gc_client.open_by_key('1l2jZGLFoqxta2jz1RRJDTa1x15nHmXxGa2ZXhWXjM4M')
ws = sh.worksheet("PIVOT")
rows = ws.get_all_values()
print("PIVOT rows 1 to 30:")
for idx in range(min(30, len(rows))):
    print(f"Row {idx+1}: {rows[idx]}")
