# -*- coding: utf-8 -*-
import sys
import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')

SHEET_ID = "1MtbZBgRFwCWj6uQKsSqddiJ2GsTiEvKxRIPSshDa5PM"
CREDENTIALS_FILE = "credentials.json"

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

scope = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]
creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
client = gspread.authorize(creds)
spreadsheet = client.open_by_key(SHEET_ID)

rec_vals = spreadsheet.worksheet("báo cáo tuyển dụng").get_all_values()
df_rec = pd.DataFrame(rec_vals)

data_vals = spreadsheet.worksheet("data").get_all_values()
df_data = pd.DataFrame(data_vals[1:], columns=data_vals[0])

tn_vals = spreadsheet.worksheet("thu nhập").get_all_values()
df_tn = pd.DataFrame(tn_vals[1:], columns=tn_vals[0])
df_tn['don_gan_num'] = df_tn['Số đơn gán Giao'].apply(clean_num)
df_tn['gtc_num'] = df_tn['Đơn giao tính lương'].apply(clean_num)
df_tn['luong_num'] = df_tn['Tổng lương'].apply(clean_num)
df_tn['Emp_Code'] = df_tn['Nhân viên'].apply(lambda x: str(x).split('-')[0].strip())

ns_vals = spreadsheet.worksheet("năng suất").get_all_values()
df_ns = pd.DataFrame(ns_vals[1:], columns=ns_vals[0])
df_ns['TongDon_num'] = df_ns['TongDon'].apply(clean_num)
df_ns['TongDonGTC_num'] = df_ns['TongDonGTC'].apply(clean_num)
df_ns['%GTC_num'] = df_ns['%GTC'].str.replace('%', '').str.replace(',', '.').apply(clean_num)
df_ns['Emp_Code'] = df_ns['NhanVien'].apply(lambda x: str(x).split('_')[0].strip())

hubs = ["22830000", "21377000"]

for hub_query in hubs:
    print(f"\n=================== HUB: {hub_query} ===================")
    rec_row = None
    for idx, r in df_rec.iterrows():
        row_str = " ".join([str(x) for x in r])
        if hub_query in row_str:
            rec_row = r
            break
    
    code = rec_row[2] if rec_row is not None else hub_query
    bc_short_name = rec_row[3] if rec_row is not None else hub_query
    full_bc_name = rec_row[4] if rec_row is not None else hub_query
    main_name = bc_short_name.split(')')[-1].strip() if ')' in bc_short_name else hub_query

    print("Code:", code, "| Short name:", bc_short_name, "| Main name:", main_name)

    # Thu Nhập
    sub_tn = df_tn[
        df_tn['Bưu cục'].str.contains(code, regex=False, na=False) |
        df_tn['Bưu cục'].str.contains(bc_short_name, regex=False, na=False) |
        df_tn['Bưu cục'].str.lower().str.contains(main_name.lower(), regex=False, na=False)
    ]
    print(f"Thu Nhập rows ({len(sub_tn)}):")
    avg_gan = sub_tn['don_gan_num'].mean() if len(sub_tn) > 0 else 0
    avg_gtc = sub_tn['gtc_num'].mean() if len(sub_tn) > 0 else 0
    print(f"  Avg Gán: {avg_gan:.1f}, Avg GTC: {avg_gtc:.1f}")
    
    for tn_cat in ['Dưới 6 tháng', '6 tháng - 3 năm', 'Trên 3 năm']:
        sub_cat = sub_tn[sub_tn['Thâm niên'].str.contains(tn_cat, regex=False, na=False)]
        mean_luong = sub_cat['luong_num'].mean() if len(sub_cat) > 0 else 0
        print(f"  Lương '{tn_cat}' (cnt={len(sub_cat)}): {mean_luong:,.0f} đ")

    # Năng Suất
    sub_ns_all = df_ns[
        df_ns['Bưu cục'].str.contains(bc_short_name, regex=False, na=False) |
        df_ns['Bưu cục'].str.lower().str.contains(main_name.lower(), regex=False, na=False)
    ]
    print(f"Năng Suất total rows for hub: {len(sub_ns_all)}")
    print("Available dates in NS for this hub:", sub_ns_all['Ngay'].unique())

    # Try filtering for '24' or latest date
    sub_ns_24 = sub_ns_all[sub_ns_all['Ngay'].str.contains('24')]
    print(f"NS rows for date '24' (cnt={len(sub_ns_24)}):")
    if len(sub_ns_24) > 0:
        merged_24 = pd.merge(sub_ns_24, df_tn[['Emp_Code', 'Thâm niên', 'Ngày vào làm']], on='Emp_Code', how='left')
        print(f"  Merged NS rows: {len(merged_24)}")
        good = len(merged_24[merged_24['TongDonGTC_num'] >= 50])
        mid = len(merged_24[(merged_24['TongDonGTC_num'] >= 35) & (merged_24['TongDonGTC_num'] < 50)])
        low = len(merged_24[merged_24['TongDonGTC_num'] < 35])
        print(f"  >50 GTC: {good}, 35-50 GTC: {mid}, <35 GTC: {low}")
