import os
import sys
import io
import pandas as pd
import numpy as np
import gspread
from google.oauth2.service_account import Credentials

os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except AttributeError:
    pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PARENT_DIR = os.path.dirname(BASE_DIR)
JSON_FILE = os.path.join(PARENT_DIR, 'credentials.json')
SHEET_KEY = '1JZ1eRerRqrpwjZ4HBevQunjd8VquM_cvPFz12TaJfMQ'

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def main():
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    sh = gc_client.open_by_key(SHEET_KEY)
    
    # ----------------------------------------------------
    # 1. Sheet trên10kg (Sản lượng Tạo)
    # ----------------------------------------------------
    ws1 = sh.worksheet('trên10kg')
    data1 = ws1.get_all_records()
    df1 = pd.DataFrame(data1)
    df1['so_don'] = pd.to_numeric(df1['so_don'], errors='coerce').fillna(0)
    df1['hen_lay'] = pd.to_datetime(df1['hen_lay'], errors='coerce')
    
    # Filter valid dates up to 2026-08-05 (2026-08-06 and 07 are partial/incomplete)
    df1_clean = df1[df1['hen_lay'] >= '2026-07-19'].copy()
    
    daily_created = df1_clean.groupby(df1_clean['hen_lay'].dt.strftime('%Y-%m-%d'))['so_don'].sum().reset_index()
    daily_created.columns = ['Date', 'Created_Orders']
    
    # Customer Group Breakdown
    kh_breakdown = df1_clean.groupby('nhom_kh')['so_don'].agg(['sum', 'count']).reset_index()
    kh_breakdown['pct'] = (kh_breakdown['sum'] / kh_breakdown['sum'].sum()) * 100
    
    # Weight Group Breakdown
    kl_breakdown = df1_clean.groupby('nhom_kl')['so_don'].agg(['sum', 'count']).reset_index()
    kl_breakdown['pct'] = (kl_breakdown['sum'] / kl_breakdown['sum'].sum()) * 100
    
    # Province Breakdown
    prov_created = df1_clean.groupby('province_name')['so_don'].sum().reset_index().sort_values(by='so_don', ascending=False)
    
    # Top 10 Bưu cục (Warehouse) Created
    bc_created = df1_clean.groupby(['warehouse_name', 'province_name'])['so_don'].sum().reset_index().sort_values(by='so_don', ascending=False)
    
    # ----------------------------------------------------
    # 2. Sheet SL > 10kg (Sản lượng Về & Vận Hành)
    # ----------------------------------------------------
    ws2 = sh.worksheet('SL > 10kg')
    data2 = ws2.get_all_records()
    df2 = pd.DataFrame(data2)
    df2['Volume_num'] = pd.to_numeric(df2['Volume'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
    df2['GTC_num'] = pd.to_numeric(df2['Sản Lượng Giao Thành Công'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
    df2['Ton_num'] = pd.to_numeric(df2['Sản Lượng Tồn'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
    df2['HangMoi_num'] = pd.to_numeric(df2['Hàng Mới Về Trong Ngày'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
    
    # Extract clean YYYY-MM-DD from 'Time' column (e.g., '2026-07-24 - Thứ 6')
    df2['Date'] = df2['Time'].str.extract(r'(\d{4}-\d{2}-\d{2})')
    
    daily_arr = df2.groupby('Date').agg({
        'Volume_num': 'sum',
        'HangMoi_num': 'sum',
        'GTC_num': 'sum',
        'Ton_num': 'sum'
    }).reset_index().sort_values(by='Date')
    
    # Merged Daily Comparison
    merged_daily = pd.merge(daily_created, daily_arr, on='Date', how='outer').fillna(0)
    merged_daily['DoD_Created_%'] = merged_daily['Created_Orders'].pct_change() * 100
    merged_daily['DoD_Volume_%'] = merged_daily['Volume_num'].pct_change() * 100
    
    # Print out summary report data
    print("\n--- 1. BẢNG TỔNG HỢP THEO NGÀY (SẢN LƯỢNG TẠO VS SẢN LƯỢNG VỀ) ---")
    print(merged_daily.to_string(index=False))
    
    print("\n--- 2. PHÂN NHÓM KHÁCH HÀNG (SẢN LƯỢNG TẠO) ---")
    print(kh_breakdown.to_string(index=False))
    
    print("\n--- 3. PHÂN NHÓM KHỐI LƯỢNG (SẢN LƯỢNG TẠO) ---")
    print(kl_breakdown.to_string(index=False))
    
    print("\n--- 4. PHÂN THEO TỈNH (SẢN LƯỢNG TẠO) ---")
    print(prov_created.to_string(index=False))
    
    print("\n--- 5. TOP 10 BƯU CỤC TẠO SẢN LƯỢNG CAO NHẤT ---")
    print(bc_created.head(10).to_string(index=False))
    
    # Top 10 Bưu cục Về Volume
    bc_arr = df2.groupby(['Chi tiết', 'Tỉnh', 'AM'])['Volume_num'].sum().reset_index().sort_values(by='Volume_num', ascending=False)
    print("\n--- 6. TOP 10 BƯU CỤC CÓ SẢN LƯỢNG VỀ (VOLUME) CAO NHẤT ---")
    print(bc_arr.head(10).to_string(index=False))

    # Spike Detection (Ngày đột biến)
    mean_created = merged_daily[merged_daily['Created_Orders'] > 100]['Created_Orders'].mean()
    std_created = merged_daily[merged_daily['Created_Orders'] > 100]['Created_Orders'].std()
    spikes = merged_daily[merged_daily['Created_Orders'] > (mean_created + std_created)]
    print(f"\n--- 7. CÁC NGÀY ĐỘT BIẾN TẠO ĐƠN (Trung bình = {mean_created:.1f}, Std = {std_created:.1f}) ---")
    print(spikes[['Date', 'Created_Orders', 'DoD_Created_%']].to_string(index=False))

if __name__ == "__main__":
    main()
