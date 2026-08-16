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
rec_vals = ws_rec.get_all_values()

print("=== ALL ROWS IN BÁO CÁO TUYỂN DỤNG ===")
for r_idx in range(min(5, len(rec_vals))):
    print(f"Row {r_idx}: {rec_vals[r_idx]}")

targets = ['20942000', '22830000', 'Di Linh', 'Cam Linh']

print("\n=== TARGET ROWS ===")
for idx, r in enumerate(rec_vals):
    r_str = " | ".join([str(x) for x in r])
    if any(t in r_str for t in targets):
        print(f"\nRow {idx}:")
        for col_idx, val in enumerate(r):
            if val:
                header = rec_vals[1][col_idx] if len(rec_vals) > 1 else f"Col{col_idx}"
                header0 = rec_vals[0][col_idx] if len(rec_vals) > 0 else ""
                print(f"  Col {col_idx} [{header0} / {header}]: {val}")

