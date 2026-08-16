import sys
import gspread
from oauth2client.service_account import ServiceAccountCredentials

sys.stdout.reconfigure(encoding='utf-8')

SCOPE = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', SCOPE)
client = gspread.authorize(creds)

sheet_id = '1MtbZBgRFwCWj6uQKsSqddiJ2GsTiEvKxRIPSshDa5PM'
spreadsheet = client.open_by_key(sheet_id)

ws_rec = spreadsheet.worksheet("báo cáo tuyển dụng")
vals = ws_rec.get_all_values()

for r in range(min(5, len(vals))):
    print(f"\n--- ROW {r} ---")
    row_vals = vals[r]
    for c in range(len(row_vals)):
        print(f"Col {c:2d}: '{row_vals[c]}'")

