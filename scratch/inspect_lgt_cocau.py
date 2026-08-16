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

targets = ['20942000', '22830000', 'Di Linh', 'Cam Linh']

# Let's check LấyGiaoTrảMới
ws = spreadsheet.worksheet("LấyGiaoTrảMới")
print(f"=== LấyGiaoTrảMới (rows: {ws.row_count}, cols: {ws.col_count}) ===")
# read first row for headers
headers = ws.row_values(1)
print("Headers:", headers)

# Search for targets in LấyGiaoTrảMới
all_vals = ws.get_all_values()
df_lgt = pd.DataFrame(all_vals[1:], columns=all_vals[0])
for t in targets:
    cnt = df_lgt.apply(lambda r: r.astype(str).str.contains(t).any(), axis=1).sum()
    print(f"Target '{t}' count in LấyGiaoTrảMới: {cnt}")

# Let's inspect CơCấuVùng sheet completely
ws_cc = spreadsheet.worksheet("CơCấuVùng")
print("\n=== SHEET: CơCấuVùng ===")
cc_vals = ws_cc.get_all_values()
for idx, r in enumerate(cc_vals):
    r_str = " | ".join([str(x) for x in r if str(x).strip()])
    if any(t in r_str for t in targets):
        print(f"Row {idx+1}: {r_str}")

