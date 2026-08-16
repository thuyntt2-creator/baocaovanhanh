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

ws_rec = spreadsheet.worksheet("báo cáo tuyển dụng")
vals = ws_rec.get_all_values()

print("=== CHECKING COL 7 (Trạng thái), COL 8, COL 9, COL 10 Across Sample Rows ===")
for r_idx in range(2, min(20, len(vals))):
    r = vals[r_idx]
    bc_name = r[3]
    status_pttt = r[7]
    col8 = r[8]
    col9 = r[9]
    col10 = r[10]
    status_xl = r[18] if len(r) > 18 else ""
    col19 = r[19] if len(r) > 19 else ""
    col20 = r[20] if len(r) > 20 else ""
    col21 = r[21] if len(r) > 21 else ""
    print(f"Row {r_idx:2d} | {bc_name:20s} | PTTT: {status_pttt:5s} | Col8={col8:2s}, Col9={col9:2s}, Col10={col10:2s} | XL: {status_xl:5s} | Col19={col19:2s}, Col20={col20:2s}, Col21={col21:2s}")

