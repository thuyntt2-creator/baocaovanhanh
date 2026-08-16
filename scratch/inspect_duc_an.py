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
df_rec = pd.DataFrame(ws_rec.get_all_values())

ws_data = spreadsheet.worksheet("data")
data_vals = ws_data.get_all_values()
df_data = pd.DataFrame(data_vals[1:], columns=data_vals[0])

ws_tn = spreadsheet.worksheet("thu nhập")
tn_vals = ws_tn.get_all_values()
df_tn = pd.DataFrame(tn_vals[1:], columns=tn_vals[0])

ws_ns = spreadsheet.worksheet("năng suất")
ns_vals = ws_ns.get_all_values()
df_ns = pd.DataFrame(ns_vals[1:], columns=ns_vals[0])

query = '21477000'
query_name = 'Đức An'

print(f"=== RECRUITMENT SHEET MATCHES FOR '{query}' / '{query_name}' ===")
for r_idx, r in df_rec.iterrows():
    r_str = " ".join([str(x) for x in r])
    if query in r_str or query_name in r_str:
        print(f"Row {r_idx}: ID='{r[2]}', Bưu cục='{r[3]}', BC='{r[4]}'")

print(f"\n=== DATA SHEET MATCHES FOR '{query}' / '{query_name}' ===")
sub_d = df_data[df_data.apply(lambda r: r.astype(str).str.contains(query_name).any(), axis=1)]
print("Matches in Data:", len(sub_d))
if len(sub_d) > 0:
    print("Unique Chi tiết names:", sub_d['Chi tiết'].unique())
    print("Unique dates in Data for Đức An:", sub_d['Time'].unique())
    latest_date = sub_d['Time'].max()
    print("Latest date:", latest_date)
    sub_latest = sub_d[sub_d['Time'] == latest_date]
    print(sub_latest[['Chi tiết', 'Loại Hàng', 'Time', 'Volume', 'Sản Lượng Giao Thành Công', 'Sản Lượng Gán']].to_string())

print(f"\n=== THU NHẬP SHEET MATCHES FOR '{query}' / '{query_name}' ===")
sub_tn = df_tn[df_tn.apply(lambda r: r.astype(str).str.contains(query_name).any(), axis=1)]
print("Matches in Thu Nhập:", len(sub_tn))
if len(sub_tn) > 0:
    print(sub_tn[['Nhân viên', 'Bưu cục', 'Thâm niên', 'Tổng lương']].to_string())

print(f"\n=== NĂNG SUẤT SHEET MATCHES FOR '{query}' / '{query_name}' ===")
sub_ns = df_ns[df_ns.apply(lambda r: r.astype(str).str.contains(query_name).any(), axis=1)]
print("Matches in Năng Suất:", len(sub_ns))
if len(sub_ns) > 0:
    print("Unique Bưu cục in Năng suất:", sub_ns['Bưu cục'].unique())
    print("Unique dates in Năng suất:", sub_ns['Ngay'].unique())

