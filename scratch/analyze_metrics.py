import sys
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import numpy as np

sys.stdout.reconfigure(encoding='utf-8')

SCOPE = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', SCOPE)
client = gspread.authorize(creds)

sheet_id = '1MtbZBgRFwCWj6uQKsSqddiJ2GsTiEvKxRIPSshDa5PM'
spreadsheet = client.open_by_key(sheet_id)

print("=== GỬI SẾP SHEET FULL CONTENT ===")
ws_gs = spreadsheet.worksheet("Gửi sếp")
for i, r in enumerate(ws_gs.get_all_values()):
    if any(r):
        print(f"L{i+1}: {r}")

print("\n=== DATA SHEET LATEST DATES ===")
ws_data = spreadsheet.worksheet("data")
data_vals = ws_data.get_all_values()
df_data = pd.DataFrame(data_vals[1:], columns=data_vals[0])

for name in ['Di Linh', 'Cam Linh']:
    sub = df_data[df_data['Chi tiết'].str.contains(name, na=False)]
    print(f"\n--- Data for {name} ---")
    print("Latest dates:", sub['Time'].unique())
    latest_date = sub['Time'].max()
    sub_latest = sub[sub['Time'] == latest_date]
    print(f"Date: {latest_date}")
    print(sub_latest[['Chi tiết', 'Time', 'Volume', '% Gán', '% GTC', 'Leadtime', 'Sản Lượng Giao Thành Công', 'Sản Lượng Gán', 'Sản Lượng Tồn', 'Sản Lượng Chưa Gán']].to_string())

print("\n=== THU NHẬP SHEET CALCULATIONS ===")
ws_tn = spreadsheet.worksheet("thu nhập")
tn_vals = ws_tn.get_all_values()
df_tn = pd.DataFrame(tn_vals[1:], columns=tn_vals[0])

# Clean numeric columns
def clean_num(val):
    if not val:
        return 0.0
    val_str = str(val).replace('.', '').replace(',', '.')
    try:
        return float(val_str)
    except:
        return 0.0

df_tn['don_gan_clean'] = df_tn['Số đơn gán Giao'].apply(clean_num)
df_tn['gtc_clean'] = df_tn['Đơn giao tính lương'].apply(clean_num)
df_tn['luong_clean'] = df_tn['Tổng lương'].apply(clean_num)

for code, name in [('20942000', 'Di Linh'), ('22830000', 'Cam Linh')]:
    sub = df_tn[df_tn['Bưu cục'].str.contains(code, na=False)]
    print(f"\n==================== {code} - {name} ====================")
    print("Total Shippers:", len(sub))
    avg_gan = sub['don_gan_clean'].mean()
    avg_gtc = sub['gtc_clean'].mean()
    print(f"Năng suất TB (Đơn gán/NV): {avg_gan:.1f} (or rounded {round(avg_gan)})")
    print(f"GTC/NV: {avg_gtc:.1f} (or rounded {round(avg_gtc)})")
    
    print("\nThu nhập theo thâm niên:")
    for group, group_df in sub.groupby('Thâm niên'):
        avg_luong = group_df['luong_clean'].mean()
        count = len(group_df)
        print(f"  + {group} ({count} NV): {avg_luong:,.0f} đ".replace(',', '.'))

