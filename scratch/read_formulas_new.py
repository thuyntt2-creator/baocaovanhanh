import json
import sys
import io
import gspread
from google.oauth2.service_account import Credentials

try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except Exception:
    pass

JSON_FILE = r"c:\Users\lap4all\Documents\Auto report\credentials.json"
scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
gc_client = gspread.authorize(creds)

new_key = '1vCxSTNgSpO9ETvVRElGyuGc7lnx7LxLRhAB4-lJMHLU'
sh = gc_client.open_by_key(new_key)
ws = sh.worksheet("Đơn giao aging trên 5 ngày")

print("Formulas in new sheet Đơn giao aging trên 5 ngày:")
try:
    data = ws.get_values(value_render_option="FORMULA")
    for r_idx in range(min(5, len(data))):
        print(f"Row {r_idx+1}: {data[r_idx][:15]}")
except Exception as e:
    print("Error:", e)
