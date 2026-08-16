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

ws_ns = spreadsheet.worksheet("năng suất")
vals = ws_ns.get_all_values()

df_ns = pd.DataFrame(vals[1:], columns=vals[0])

print("Total rows in năng suất:", len(df_ns))
print("Available unique dates in sheet (sample):", df_ns['Ngay'].unique()[:15])

targets = ['Di Linh', 'Cam Linh', '20942000', '22830000']

for target in targets:
    sub = df_ns[df_ns['Bưu cục'].str.contains(target, na=False) | df_ns['NhanVien'].str.contains(target, na=False)]
    print(f"\n==================== Target: {target} (Found {len(sub)} rows) ====================")
    if len(sub) > 0:
        print("Unique Bưu cục names found:", sub['Bưu cục'].unique())
        print("Unique dates for this hub:", sub['Ngay'].unique())

