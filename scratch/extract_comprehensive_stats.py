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
tn_vals = ws_tn.get_all_values()
df_tn = pd.DataFrame(tn_vals[1:], columns=tn_vals[0])

ws_data = spreadsheet.worksheet("data")
data_vals = ws_data.get_all_values()
df_data = pd.DataFrame(data_vals[1:], columns=data_vals[0])

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

hubs = [
    ('20942000', '(LDO) Di Linh'),
    ('22830000', '(KHO) Cam Linh')
]

for code, name in hubs:
    print(f"\n=======================================================")
    print(f"BƯU CỤC {code} - {name}")
    print(f"=======================================================")
    
    # 1. DATA SHEET STATS (Latest date 2026-07-22)
    sub_d = df_data[df_data['Chi tiết'].str.contains(code, na=False) | df_data['Chi tiết'].str.contains(name.split(')')[-1].strip(), na=False)]
    latest_date = sub_d['Time'].max()
    sub_d_latest = sub_d[sub_d['Time'] == latest_date]
    print(f"--- Data Sheet ({latest_date}) ---")
    for idx, r in sub_d_latest.iterrows():
        print(f"Loại hàng: {r['Loại Hàng']} | Volume: {r['Volume']} | % Gán: {r['% Gán']} | % GTC: {r['% GTC']} | GTC: {r['Sản Lượng Giao Thành Công']} | Gán: {r['Sản Lượng Gán']} | Tồn: {r['Sản Lượng Tồn']} | Chưa gán: {r['Sản Lượng Chưa Gán']}")

    # 2. THU NHẬP SHEET STATS
    sub_t = df_tn[df_tn['Bưu cục'].str.contains(code, na=False)]
    print(f"\n--- Thu Nhập Sheet ({len(sub_t)} NV) ---")
    
    # Seniority counts
    nv_moi = sub_t[sub_t['Thâm niên'].str.contains('Dưới 6 tháng', na=False)]
    nv_cu_mid = sub_t[sub_t['Thâm niên'].str.contains('6 tháng - 3 năm', na=False)]
    nv_cu_old = sub_t[sub_t['Thâm niên'].str.contains('Trên 3 năm', na=False)]
    
    print(f"Phân loại nhân sự:")
    print(f"  - Dưới 6 tháng (Mới): {len(nv_moi)} NV")
    print(f"  - Trên 6 tháng - 3 năm: {len(nv_cu_mid)} NV")
    print(f"  - Trên 3 năm: {len(nv_cu_old)} NV")
    print(f"  -> Tổng cũ (>= 6 tháng): {len(nv_cu_mid) + len(nv_cu_old)} NV")
    
    avg_gan = sub_t['don_gan_clean'].mean()
    avg_gtc = sub_t['gtc_clean'].mean()
    print(f"\nNăng suất TB:")
    print(f"  - Số đơn gán/NV: {avg_gan:.0f} đơn (chính xác: {avg_gan:.1f})")
    print(f"  - GTC/NV: {avg_gtc:.0f} đơn (chính xác: {avg_gtc:.1f})")
    
    print(f"\nThu nhập cụ thể (Trung bình):")
    if len(nv_moi) > 0:
        print(f"  + NV Dưới 6 tháng : {nv_moi['luong_clean'].mean():,.0f} đ".replace(',', '.'))
    if len(nv_cu_mid) > 0:
        print(f"  + NV Trên 6 tháng - 3 năm : {nv_cu_mid['luong_clean'].mean():,.0f} đ".replace(',', '.'))
    if len(nv_cu_old) > 0:
        print(f"  + NV Trên 3 năm : {nv_cu_old['luong_clean'].mean():,.0f} đ".replace(',', '.'))
        
    print(f"\nDanh sách NV chi tiết:")
    for idx, r in sub_t.iterrows():
        print(f"  - {r['Nhân viên']} | Thâm niên: {r['Thâm niên']} | Đơn gán: {r['don_gan_clean']} | GTC: {r['gtc_clean']} | Lương: {r['luong_clean']:,.0f}đ".replace(',', '.'))

