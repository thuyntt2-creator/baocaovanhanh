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

orig_sheet_id = '14eU6peb0rb3Bqp7Q4l4LkeDKVlrbcU0Qp-u1R2wjbQw'
orig_ss = client.open_by_key(orig_sheet_id)

ws_rec = orig_ss.worksheet("báo cáo tuyển dụng")
vals = ws_rec.get_all_values()

print("=== ORIGINAL SHEET HEADERS (First 5 rows) ===")
for r_idx in range(min(5, len(vals))):
    print(f"\nRow {r_idx}:")
    for c_idx, val in enumerate(vals[r_idx]):
        if val != "":
            print(f"  Col {c_idx:2d}: {val}")

