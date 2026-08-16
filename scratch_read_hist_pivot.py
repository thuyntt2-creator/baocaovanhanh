import json
import sys
import io
import gspread
from google.oauth2.service_account import Credentials

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

BASE_DIR = r"c:\Users\lap4all\Documents\Auto report"
JSON_FILE = r"c:\Users\lap4all\Documents\Auto report\credentials.json"
scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
gc_client = gspread.authorize(creds)

for key, name in [('1WCzgao34cA_SttyB9ytHfE1qKTNl_3iFqDbEfw3lbyU', '1WCz (Main)'), ('1l2jZGLFoqxta2jz1RRJDTa1x15nHmXxGa2ZXhWXjM4M', '1l2j (Follow Gan)')]:
    print(f"\n=== {name} PIVOT rows 31-51 ===")
    try:
        sh = gc_client.open_by_key(key)
        ws = sh.worksheet("PIVOT")
        rows = ws.get_all_values()
        for idx in range(30, min(52, len(rows))):
            print(f"Row {idx+1}: {rows[idx]}")
    except Exception as e:
        print("Error:", e)
