import sys
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')

SCOPE = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', SCOPE)
client = gspread.authorize(creds)

sheet_id = '1MtbZBgRFwCWj6uQKsSqddiJ2GsTiEvKxRIPSshDa5PM'
spreadsheet = client.open_by_key(sheet_id)

print("Worksheets:")
for ws in spreadsheet.worksheets():
    print(f"Sheet: {ws.title} | id: {ws.id}")

print("\n=== INSPECTING SHEET: bưu cục ===")
try:
    ws_bc = spreadsheet.worksheet("bưu cục")
    vals = ws_bc.get_all_values()
    print("Values in sheet 'bưu cục':")
    for r_idx, r in enumerate(vals):
        print(f"Row {r_idx}: {r}")
except Exception as e:
    print("Error opening sheet 'bưu cục':", e)

