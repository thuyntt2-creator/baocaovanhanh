import sys
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials

sys.stdout.reconfigure(encoding='utf-8')

SHEET_ID = "1MtbZBgRFwCWj6uQKsSqddiJ2GsTiEvKxRIPSshDa5PM"
CREDENTIALS_FILE = "credentials.json"

scope = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]
creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
client = gspread.authorize(creds)
spreadsheet = client.open_by_key(SHEET_ID)

ws_hotro = spreadsheet.worksheet("NV HỖ TRỢ")
vals_hotro = ws_hotro.get_all_values()
print("Raw rows in 'NV HỖ TRỢ':")
for idx, r in enumerate(vals_hotro):
    print(f"Row {idx}: {r}")

tn_vals = spreadsheet.worksheet("thu nhập").get_all_values()
df_tn = pd.DataFrame(tn_vals[1:], columns=tn_vals[0])
sub_tn_camlinh = df_tn[df_tn['Bưu cục'].str.contains("22830000", na=False) | df_tn['Bưu cục'].str.contains("Cam Linh", na=False)]

print("\nAll 16 employees in 'thu nhập' for Cam Linh:")
for idx, r in sub_tn_camlinh.iterrows():
    print(f" - {r['Nhân viên']} | Thâm niên: {r['Thâm niên']} | Lương: {r['Lương HH/ ngày']}")
