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

# Clean numeric columns
def clean_num(val):
    if not val:
        return 0
    val_str = str(val).replace('.', '').replace(',', '.')
    try:
        return float(val_str)
    except:
        return 0

df_ns['TongDon_num'] = df_ns['TongDon'].apply(clean_num)
df_ns['TongDonGTC_num'] = df_ns['TongDonGTC'].apply(clean_num)

hubs = [
    ('(LDO) Di Linh', '20942000 - (LDO) Di Linh'),
    ('(KHO) Cam Linh', '22830000 - (KHO) Cam Linh')
]

dates = ['22 thg 7, 2026', '21 thg 7, 2026']

for hub_name, full_name in hubs:
    print(f"\n=======================================================")
    print(f"BƯU CỤC: {full_name}")
    print(f"=======================================================")
    
    sub_hub = df_ns[df_ns['Bưu cục'] == hub_name]
    
    for date in dates:
        sub_date = sub_hub[sub_hub['Ngay'] == date]
        print(f"\n--- NGÀY: {date} (Số NV đi làm: {len(sub_date)}) ---")
        if len(sub_date) == 0:
            print("Không có dữ liệu cho ngày này.")
            continue
        
        # Sort by TongDonGTC descending
        sub_date_sorted = sub_date.sort_values(by='TongDonGTC_num', ascending=False)
        
        tot_don = sub_date_sorted['TongDon_num'].sum()
        tot_gtc = sub_date_sorted['TongDonGTC_num'].sum()
        avg_don = sub_date_sorted['TongDon_num'].mean()
        avg_gtc = sub_date_sorted['TongDonGTC_num'].mean()
        overall_gtc_rate = (tot_gtc / tot_don * 100) if tot_don > 0 else 0
        
        print(f"TỔNG QUAN HÔM QUA ({date}):")
        print(f"  - Tổng số NV đi làm: {len(sub_date)} NV")
        print(f"  - Tổng đơn gán: {tot_don:.0f} đơn | TB/NV: {avg_don:.1f} đơn")
        print(f"  - Tổng đơn GTC: {tot_gtc:.0f} đơn | TB GTC/NV: {avg_gtc:.1f} đơn")
        print(f"  - Tỷ lệ GTC toàn bưu cục: {overall_gtc_rate:.2f}%\n")
        
        print("CHI TIẾT NĂNG SUẤT TỪNG NHÂN VIÊN:")
        for idx, r in sub_date_sorted.iterrows():
            nv = r['NhanVien']
            tong = r['TongDon']
            gtc = r['TongDonGTC']
            rate = r['%GTC']
            lt = r['LT']
            chuyen = r['ChuyenDi']
            print(f"  + {nv:35s} | Đơn gán: {tong:4s} | GTC: {gtc:4s} | %GTC: {rate:7s} | LT: {lt:5s} | Chuyến đi: {chuyen}")

