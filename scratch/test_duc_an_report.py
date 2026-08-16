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

def clean_num(val):
    if not val or pd.isna(val):
        return 0.0
    s = str(val).strip().replace('đ', '').replace('%', '').strip()
    if ',' in s and '.' in s:
        if s.rfind(',') > s.rfind('.'):
            s = s.replace('.', '').replace(',', '.')
        else:
            s = s.replace(',', '')
    elif ',' in s:
        parts = s.split(',')
        if len(parts) == 2 and len(parts[1]) == 3 and not parts[1].endswith('00'):
            s = s.replace(',', '')
        else:
            s = s.replace(',', '.')
    elif '.' in s:
        parts = s.split('.')
        if len(parts) > 1 and all(len(p) == 3 for p in parts[1:]):
            s = s.replace('.', '')
    try:
        return float(s)
    except ValueError:
        return 0.0

ws_data = spreadsheet.worksheet("data")
df_data = pd.DataFrame(ws_data.get_all_values()[1:], columns=ws_data.get_all_values()[0])

ws_tn = spreadsheet.worksheet("thu nhập")
df_tn = pd.DataFrame(ws_tn.get_all_values()[1:], columns=ws_tn.get_all_values()[0])
df_tn['don_gan_num'] = df_tn['Số đơn gán Giao'].apply(clean_num)
df_tn['gtc_num'] = df_tn['Đơn giao tính lương'].apply(clean_num)
df_tn['luong_num'] = df_tn['Tổng lương'].apply(clean_num)
df_tn['Emp_Code'] = df_tn['Nhân viên'].apply(lambda x: str(x).split('-')[0].strip())

ws_ns = spreadsheet.worksheet("năng suất")
df_ns = pd.DataFrame(ws_ns.get_all_values()[1:], columns=ws_ns.get_all_values()[0])
df_ns['TongDon_num'] = df_ns['TongDon'].apply(clean_num)
df_ns['TongDonGTC_num'] = df_ns['TongDonGTC'].apply(clean_num)
df_ns['%GTC_num'] = df_ns['%GTC'].str.replace('%', '').str.replace(',', '.').apply(clean_num)
df_ns['Emp_Code'] = df_ns['NhanVien'].apply(lambda x: str(x).split('_')[0].strip())

query = '21477000 - (DNO) Đức An'
code = '21477000'
bc_short_name = '(DNO) Đức An'
main_name = 'Đức An'

# Matching with regex=False
sub_d = df_data[
    df_data['Chi tiết'].str.contains(code, regex=False, na=False) |
    df_data['Chi tiết'].str.contains(bc_short_name, regex=False, na=False) |
    df_data['Chi tiết'].str.lower().str.contains(main_name.lower(), regex=False, na=False)
]

print(f"Data matches for '{main_name}':", len(sub_d))
latest_date = sub_d['Time'].max()
sub_d_latest = sub_d[sub_d['Time'] == latest_date]

row_ca1 = sub_d_latest[sub_d_latest['Loại Hàng'].str.contains('Ca 1', regex=False, na=False)]
row_ton = sub_d_latest[sub_d_latest['Loại Hàng'].str.contains('Tồn', regex=False, na=False)]

vol_ca1 = clean_num(row_ca1['Volume'].values[0]) if len(row_ca1) > 0 else 0
gan_ca1 = clean_num(row_ca1['Sản Lượng Gán'].values[0]) if len(row_ca1) > 0 else 0
gtc_ca1 = clean_num(row_ca1['Sản Lượng Giao Thành Công'].values[0]) if len(row_ca1) > 0 else 0
ton_ca1 = clean_num(row_ca1['Sản Lượng Tồn'].values[0]) if len(row_ca1) > 0 else 0
chuagan_ca1 = clean_num(row_ca1['Sản Lượng Chưa Gán'].values[0]) if len(row_ca1) > 0 else 0

vol_ton = clean_num(row_ton['Volume'].values[0]) if len(row_ton) > 0 else 0
gan_ton = clean_num(row_ton['Sản Lượng Gán'].values[0]) if len(row_ton) > 0 else 0
gtc_ton = clean_num(row_ton['Sản Lượng Giao Thành Công'].values[0]) if len(row_ton) > 0 else 0
ton_ton = clean_num(row_ton['Sản Lượng Tồn'].values[0]) if len(row_ton) > 0 else 0
chuagan_ton = clean_num(row_ton['Sản Lượng Chưa Gán'].values[0]) if len(row_ton) > 0 else 0

tot_vol = vol_ca1 + vol_ton
tot_gtc = gtc_ca1 + gtc_ton
tot_gtc_rate = (tot_gtc / tot_vol * 100) if tot_vol > 0 else 0

print(f"Ca 1 Vol: {vol_ca1}, Gán: {gan_ca1}, GTC: {gtc_ca1}, Tồn: {ton_ca1}, Chưa gán: {chuagan_ca1}")
print(f"Tồn Vol: {vol_ton}, Gán: {gan_ton}, GTC: {gtc_ton}, Tồn: {ton_ton}, Chưa gán: {chuagan_ton}")
print(f"Tổng Volume: {tot_vol}, Tổng GTC: {tot_gtc} => % GTC Bưu cục: {tot_gtc_rate:.2f}%")

# Matching in Thu Nhập
sub_tn = df_tn[
    df_tn['Bưu cục'].str.contains(code, regex=False, na=False) |
    df_tn['Bưu cục'].str.contains(bc_short_name, regex=False, na=False) |
    df_tn['Bưu cục'].str.lower().str.contains(main_name.lower(), regex=False, na=False)
]

luong_g01 = sub_tn[sub_tn['Thâm niên'].str.contains('Dưới 6 tháng', regex=False, na=False)]['luong_num'].mean()
luong_g02 = sub_tn[sub_tn['Thâm niên'].str.contains('6 tháng - 3 năm', regex=False, na=False)]['luong_num'].mean()

print(f"Lương G01 (Dưới 6 tháng): {luong_g01:,.0f} đ".replace(',', '.'))
print(f"Lương G02 (6 tháng - 3 năm): {luong_g02:,.0f} đ".replace(',', '.'))

# Matching in Năng Suất
sub_ns = df_ns[
    (df_ns['Bưu cục'].str.contains(bc_short_name, regex=False, na=False) |
     df_ns['Bưu cục'].str.lower().str.contains(main_name.lower(), regex=False, na=False)) &
    (df_ns['Ngay'].str.contains('22'))
]

merged_ns = pd.merge(sub_ns, df_tn[['Emp_Code', 'Thâm niên', 'Ngày vào làm']], on='Emp_Code', how='left')
cnt_good = len(merged_ns[merged_ns['TongDonGTC_num'] >= 50])
cnt_mid = len(merged_ns[(merged_ns['TongDonGTC_num'] >= 35) & (merged_ns['TongDonGTC_num'] < 50)])
cnt_low = len(merged_ns[merged_ns['TongDonGTC_num'] < 35])

print(f"Năng suất đi làm hôm qua (Total {len(merged_ns)} NV): Good={cnt_good}, Mid={cnt_mid}, Low={cnt_low}")

