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

# Print top 3 rows with column index
for r_idx in range(min(3, len(rec_vals))):
    print(f"\n--- Row {r_idx} ---")
    for c_idx, val in enumerate(rec_vals[r_idx]):
        print(f"Col {c_idx:2d}: {val}")

# Row 10 (Di Linh) and Row 34 (Cam Linh)
for target_idx, name in [(10, "Di Linh"), (34, "Cam Linh")]:
    print(f"\n==================== Row {target_idx}: {name} ====================")
    row_data = rec_vals[target_idx]
    for c_idx in range(len(row_data)):
        h0 = rec_vals[0][c_idx] if c_idx < len(rec_vals[0]) else ""
        h1 = rec_vals[1][c_idx] if c_idx < len(rec_vals[1]) else ""
        val = row_data[c_idx]
        print(f"Col {c_idx:2d} | Header: [{h0} // {h1}] => Value: {val}")

