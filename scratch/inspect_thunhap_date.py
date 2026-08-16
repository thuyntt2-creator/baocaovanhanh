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

ws_tn = spreadsheet.worksheet("thu nhập")
vals_tn = ws_tn.get_all_values()
df_tn = pd.DataFrame(vals_tn[1:], columns=vals_tn[0])

print("Unique dates in column 'Thời gian':", df_tn['Thời gian'].unique())

for code in ['20942000', '22830000']:
    sub = df_tn[df_tn['Bưu cục'].str.contains(code, na=False)]
    print(f"\n--- Code {code} in thu nhập ---")
    print("Dates in 'Thời gian':", sub['Thời gian'].unique())
    print(sub[['Nhân viên', 'Bưu cục', 'Thời gian', 'Số đơn gán Giao', 'Đơn giao tính lương', 'Tổng lương']].head(5).to_string())

