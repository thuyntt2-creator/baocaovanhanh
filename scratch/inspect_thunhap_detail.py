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

print("=== THU NHẬP SHEET DETAILED ANALYSIS ===")
ws_tn = spreadsheet.worksheet("thu nhập")
tn_vals = ws_tn.get_all_values()
df_tn = pd.DataFrame(tn_vals[1:], columns=tn_vals[0])

for code in ['20942000', '22830000']:
    sub = df_tn[df_tn['Bưu cục'].str.contains(code, na=False)]
    print(f"\n==================== BƯU CỤC: {code} ====================")
    print("Total shippers in sheet:", len(sub))
    print(sub.to_string())

