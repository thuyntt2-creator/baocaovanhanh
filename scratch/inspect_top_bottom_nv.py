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

    print(f"\n=================== BƯU CỤC: {full_bc_name} ({code}) ===================")

    # Năng Suất latest date
    sub_ns_hub = df_ns[
        df_ns['Bưu cục'].str.contains(bc_short_name, regex=False, na=False) |
        df_ns['Bưu cục'].str.lower().str.contains(main_name.lower(), regex=False, na=False)
    ]
    
    def parse_date_key(d):
        try:
            p = str(d).replace('thg', '').replace(',', '').split()
            return (int(p[2]), int(p[1]), int(p[0]))
        except Exception:
            return (0, 0, 0)
    
    unique_dates = sorted(sub_ns_hub['Ngay'].unique(), key=parse_date_key, reverse=True)
    latest_date = unique_dates[0] if unique_dates else ""
    sub_ns_latest = sub_ns_hub[sub_ns_hub['Ngay'] == latest_date]

    sub_tn = df_tn[
        df_tn['Bưu cục'].str.contains(code, regex=False, na=False) |
        df_tn['Bưu cục'].str.contains(bc_short_name, regex=False, na=False) |
        df_tn['Bưu cục'].str.lower().str.contains(main_name.lower(), regex=False, na=False)
    ]

    merged = pd.merge(sub_ns_latest, sub_tn[['Emp_Code', 'Nhân viên', 'Thâm niên', 'luong_num', 'don_gan_num', 'gtc_num']], on='Emp_Code', how='left')

    print(f"Ngày năng suất: {latest_date} | Tổng NV: {len(merged)}")
    
    # Sort by TongDonGTC_num ascending (Bottom NV)
    sorted_bottom = merged.sort_values(by='TongDonGTC_num', ascending=True)
    print("\n--- TOP 5 NV NĂNG SUẤT KẾM (BOTTOM PFM) ---")
    for idx, r in sorted_bottom.head(5).iterrows():
        nv_name = r['NhanVien']
        tn = r['Thâm niên'] if pd.notna(r['Thâm niên']) else "N/A"
        gong = r['TongDon_num']
        gtc = r['TongDonGTC_num']
        rate = r['%GTC']
        luong = r['luong_num'] if pd.notna(r['luong_num']) else 0
        print(f"• NV: {nv_name} | Thâm niên: {tn} | Gán: {int(gong)} | GTC: {int(gtc)} ({rate}) | Lương: {luong:,.0f} đ/ngày")

    # Sort by TongDonGTC_num descending (Top NV)
    sorted_top = merged.sort_values(by='TongDonGTC_num', ascending=False)
    print("\n--- TOP 5 NV NĂNG SUẤT CAO (TOP PFM) ---")
    for idx, r in sorted_top.head(5).iterrows():
        nv_name = r['NhanVien']
        tn = r['Thâm niên'] if pd.notna(r['Thâm niên']) else "N/A"
        gong = r['TongDon_num']
        gtc = r['TongDonGTC_num']
        rate = r['%GTC']
        luong = r['luong_num'] if pd.notna(r['luong_num']) else 0
        print(f"• NV: {nv_name} | Thâm niên: {tn} | Gán: {int(gong)} | GTC: {int(gtc)} ({rate}) | Lương: {luong:,.0f} đ/ngày")
