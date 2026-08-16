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

print("=== CHECKING DATA SHEET FOR 2026-07-22 ===")
ws_data = spreadsheet.worksheet("data")
vals_data = ws_data.get_all_values()
df_data = pd.DataFrame(vals_data[1:], columns=vals_data[0])

for name in ['Di Linh', 'Cam Linh']:
    sub = df_data[(df_data['Chi tiết'].str.contains(name, na=False)) & (df_data['Time'] == '2026-07-22 - Thứ 4')]
    print(f"\n--- {name} in sheet data ---")
    print(sub[['Chi tiết', 'Loại Hàng', 'Time', 'Volume', '% Gán', '% GTC', 'Sản Lượng Giao Thành Công', 'Sản Lượng Gán']].to_string())

print("\n=== CHECKING NĂNG SUẤT SHEET FOR 22/07/2026 ===")
ws_ns = spreadsheet.worksheet("năng suất")
vals_ns = ws_ns.get_all_values()
df_ns = pd.DataFrame(vals_ns[1:], columns=vals_ns[0])

def clean_num(val):
    if not val: return 0
    try: return float(str(val).replace('.', '').replace(',', '.'))
    except: return 0

df_ns['TongDon_num'] = df_ns['TongDon'].apply(clean_num)
df_ns['TongDonGTC_num'] = df_ns['TongDonGTC'].apply(clean_num)

for hub in ['(LDO) Di Linh', '(KHO) Cam Linh']:
    sub = df_ns[(df_ns['Bưu cục'] == hub) & (df_ns['Ngay'] == '22 thg 7, 2026')]
    tot_don = sub['TongDon_num'].sum()
    tot_gtc = sub['TongDonGTC_num'].sum()
    rate = (tot_gtc / tot_don * 100) if tot_don > 0 else 0
    print(f"{hub}: Tổng Gán = {tot_don:.0f}, Tổng GTC = {tot_gtc:.0f} => Tỷ lệ = {rate:.2f}%")

