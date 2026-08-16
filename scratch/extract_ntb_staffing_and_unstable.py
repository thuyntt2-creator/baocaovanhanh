# -*- coding: utf-8 -*-
import sys, os, pandas as pd

sys.stdout.reconfigure(encoding='utf-8')

ntb_excel_path = r'C:\Users\lap4all\Downloads\config_psbba_NTB.xlsx'

xl = pd.ExcelFile(ntb_excel_path)

print('=== SHEET NHÂN SỰ ===')
df_ns = pd.read_excel(xl, sheet_name='Nhân sự')
print(f'Total rows in Sheet Nhân sự: {len(df_ns)}')
print('Columns:', df_ns.columns.tolist()[:15])

# Print counts by Trạng thái
if 'Trạng thái' in df_ns.columns:
    print('\nBreakdown by Trạng thái (NVPTTT):')
    print(df_ns['Trạng thái'].value_counts(dropna=False))

# Filter Unstable/Warning BCs in Sheet Nhân sự
target_ns_cols = ['Tỉnh', 'AM', 'Bưu cục', 'Trạng thái', 'Số NVPTTT thiếu đầu tuần 30/2026', 'Số NVPTTT hiện hữu', 'Định biên', 'Số NVPTTT còn thiếu đến hiện tại', 'Số NVXL thiếu tuần 30/2026', 'Số NVXL hiện hữu', 'Định biên.1']
cols_avail = [c for c in target_ns_cols if c in df_ns.columns]

print('\n--- ALL BCs in Sheet Nhân sự ---')
for idx, r in df_ns[cols_avail].iterrows():
    bc_name = r.get('Bưu cục', '')
    status = r.get('Trạng thái', '')
    thieu = r.get('Số NVPTTT thiếu đầu tuần 30/2026', 0)
    hien_huu = r.get('Số NVPTTT hiện hữu', 0)
    dinh_bien = r.get('Định biên', 0)
    thieu_ht = r.get('Số NVPTTT còn thiếu đến hiện tại', 0)
    am = r.get('AM', '')
    tinh = r.get('Tỉnh', '')
    print(f'{tinh:<12} | AM: {str(am):<15} | BC: {str(bc_name):<28} | TT: {str(status):<10} | Thiếu: {thieu} | HiênHuu: {hien_huu} | ĐB: {dinh_bien} | ThiếuHT: {thieu_ht}')

print('\n=== SHEET BẤT ỔN ===')
df_bo = pd.read_excel(xl, sheet_name='Bất ổn')
print(f'Total rows in Sheet Bất ổn: {len(df_bo)}')
print('Columns:', df_bo.columns.tolist())
print('Distinct Warehouses / BCs in Sheet Bất ổn:')
piv_bo = df_bo.groupby(['kho_giao_name', 'Trạng thái']).agg({
    'BL LM': 'mean',
    'BL KTC': 'mean',
    'tinh_hinh': 'first',
    'ly_do_bat_on': 'first'
}).reset_index()

print(piv_bo.to_string())
