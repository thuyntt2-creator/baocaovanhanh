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

sheet_key = '10cq3DUggZ4vXffcxweIRTRK3qiyMeWnV8gksdGwvp7s'
sh = gc_client.open_by_key(sheet_key)

ws = sh.worksheet("No attempt")
res = ws.get("H1:L5", value_render_option="FORMULA")
print("Row formulas/values for H-L:")
for idx, row in enumerate(res):
    print(f"Row {idx+1}: {row}")
