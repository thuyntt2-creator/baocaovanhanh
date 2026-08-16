# -*- coding: utf-8 -*-
import sys, os, pandas as pd

sys.stdout.reconfigure(encoding='utf-8')

ntb_excel_path = r'C:\Users\lap4all\Downloads\config_psbba_NTB.xlsx'
xl = pd.ExcelFile(ntb_excel_path)
df_ns = pd.read_excel(xl, sheet_name='Nhân sự')

print('=== EXACT HR STAFFING METRICS FOR NTB BCs ===')
# Filter out empty rows
df_ns_clean = df_ns.dropna(subset=['Bưu cục'])

target_cols = [
    'Tỉnh', 'AM', 'Bưu cục', 'Trạng thái', 
    'Số NVPTTT thiếu đầu tuần 30/2026', 'Số NVPTTT hiện hữu', 'Định biên',
    'Số NVPTTT còn thiếu đến hiện tại', 'Số NVXL thiếu tuần 30/2026', 'Số NVXL hiện hữu'
]

available_cols = [c for c in target_cols if c in df_ns_clean.columns]
print('Found columns:', available_cols)

res = []
for idx, r in df_ns_clean[available_cols].iterrows():
    bc = r.get('Bưu cục', '')
    status = r.get('Trạng thái', 'Ổn định')
    thieu_dau_tuan = r.get('Số NVPTTT thiếu đầu tuần 30/2026', 0)
    hien_huu = r.get('Số NVPTTT hiện hữu', 0)
    dinh_bien = r.get('Định biên', 0)
    thieu_hien_tai = r.get('Số NVPTTT còn thiếu đến hiện tại', 0)
    tinh = r.get('Tỉnh', '')
    am = r.get('AM', '')
    
    # Calculate deficit % if dinh_bien > 0
    thieu_val = int(thieu_dau_tuan) if pd.notna(thieu_dau_tuan) else 0
    hh_val = int(hien_huu) if pd.notna(hien_huu) else 0
    db_val = int(dinh_bien) if pd.notna(dinh_bien) else 0
    
    res.append({
        'Tỉnh': tinh,
        'AM': am,
        'Bưu cục': bc,
        'Trạng thái': status,
        'Thiếu T30': thieu_val,
        'Hiện hữu': hh_val,
        'Định biên': db_val,
        'Thiếu HT': int(thieu_hien_tai) if pd.notna(thieu_hien_tai) else 0
    })

df_res = pd.DataFrame(res)
print(f'Total Valid Post Offices: {len(df_res)}')

print('\nSummary by Trạng thái in Sheet Nhân sự:')
print(df_res['Trạng thái'].value_counts())

print('\n--- TOP BCs WITH HIGGEST STAFF DEFICITS (THIẾU NVPTTT) ---')
print(df_res.sort_values(by='Thiếu T30', ascending=False).head(20).to_string())
