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

sh = gc_client.open_by_key('1WCzgao34cA_SttyB9ytHfE1qKTNl_3iFqDbEfw3lbyU')
ws = sh.worksheet("Đơn giao aging trên 5 ngày")

print("Formulas in 1WCz Đơn giao aging trên 5 ngày:")
try:
    data = ws.get_values(value_render_option="FORMULA")
    for r_idx in range(min(5, len(data))):
        print(f"Row {r_idx+1}: {data[r_idx][:5]}")
except Exception as e:
    print("Error:", e)
