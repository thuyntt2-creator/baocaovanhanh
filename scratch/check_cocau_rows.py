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

ws_cc = spreadsheet.worksheet("CơCấuVùng")
cc_vals = ws_cc.get_all_values()
df_cc = pd.DataFrame(cc_vals)

print("CơCấuVùng matching rows:")
for idx, r in df_cc.iterrows():
    row_str = " | ".join([str(x) for x in r if str(x).strip()])
    if '20942000' in row_str or '22830000' in row_str or 'Di Linh' in row_str or 'Cam Linh' in row_str:
        print(f"Row {idx+1}: {row_str}")

