import gspread
from google.oauth2.service_account import Credentials
import sys
import io

try:
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')
except Exception:
    pass

JSON_FILE = r"c:\Users\lap4all\Documents\Auto report\credentials.json"
scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
gc_client = gspread.authorize(creds)

# Master aging sheet key
sh = gc_client.open_by_key('1WCzgao34cA_SttyB9ytHfE1qKTNl_3iFqDbEfw3lbyU')
ws = sh.worksheet("Cơ cấu")
rows = ws.get_all_values()

found = False
for idx, r in enumerate(rows):
    if "21094000" in "".join(r):
        print(f"Row {idx+1}: {r}")
        found = True

if not found:
    print("Not found 21094000 in master aging spreadsheet Cơ cấu sheet!")
