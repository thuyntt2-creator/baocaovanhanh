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

worksheets = {ws.title.lower().strip(): ws for ws in spreadsheet.worksheets()}
ws_input = None
for key in ['bưu cục', 'bưucục', 'buucuc', 'bưu_cục', 'báo cáo']:
    if key in worksheets:
        ws_input = worksheets[key]
        break

target_hubs = []
if ws_input:
    vals = ws_input.get_all_values()
    for idx, row in enumerate(vals):
        if row and row[0].strip():
            val = row[0].strip()
            if idx == 0 and val.lower() in ['bưu cục', 'tên bưu cục', 'mã bưu cục', 'stt', 'id']:
                continue
            target_hubs.append(val)

print("Target Hubs in sheet bưu cục:", target_hubs)

tn_vals = spreadsheet.worksheet("thu nhập").get_all_values()
df_tn = pd.DataFrame(tn_vals[1:], columns=tn_vals[0])

# Clean columns
df_tn['TongLuong_num'] = df_tn['Tổng lương'].apply(clean_num)
df_tn['LuongHH_num'] = df_tn['Lương HH/ ngày'].apply(clean_num)
df_tn['DonGanGiao_num'] = df_tn['Số đơn gán Giao'].apply(clean_num)
df_tn['DonGTC_num'] = df_tn['Đơn giao tính lương'].apply(clean_num)
df_tn['DonGanLay_num'] = df_tn['Số đơn gán Lấy'].apply(clean_num)
df_tn['DonLTC_num'] = df_tn['Đơn lấy tính lương'].apply(clean_num)

# Also load rec sheet to map full names if possible
rec_vals = spreadsheet.worksheet("báo cáo tuyển dụng").get_all_values()
df_rec = pd.DataFrame(rec_vals)

output_md = "# BÁO CÁO CHI TIẾT THU NHẬP VÀ NĂNG SUẤT TỪNG NHÂN VIÊN\n\n"

for hub_query in target_hubs:
    rec_match = None
    for idx, r in df_rec.iterrows():
        row_str = " ".join([str(x) for x in r])
        if hub_query.lower() in row_str.lower():
            rec_match = r
            break
            
    code = rec_match[2] if rec_match is not None else hub_query
    bc_short = rec_match[3] if rec_match is not None else hub_query
    full_bc_name = rec_match[4] if rec_match is not None else hub_query
    main_name = bc_short.split(')')[-1].strip() if ')' in bc_short else hub_query

    sub = df_tn[
        df_tn['Bưu cục'].str.contains(code, regex=False, na=False) |
        df_tn['Bưu cục'].str.contains(bc_short, regex=False, na=False) |
        df_tn['Bưu cục'].str.lower().str.contains(main_name.lower(), regex=False, na=False)
    ]
    
    output_md += f"## Bưu Cục: {full_bc_name} ({hub_query})\n"
    output_md += f"**Tổng số nhân viên:** {len(sub)}\n\n"
    
    if len(sub) == 0:
        output_md += "*Không có dữ liệu nhân viên trong sheet Thu nhập*\n\n"
        continue
        
    output_md += "| STT | Mã & Tên NV | Thâm Niên | Ngày Vào Làm | Gán Giao | GTC | Gán Lấy | LTC | Lương HH/Ngày (VNĐ) | Tổng Lương (VNĐ) |\n"
    output_md += "|---|---|---|---|---|---|---|---|---|---|\n"
    
    for i, (_, r) in enumerate(sub.iterrows(), 1):
        nv = r['Nhân viên']
        tn = r['Thâm niên']
        nvl = r['Ngày vào làm']
        gan_giao = int(r['DonGanGiao_num'])
        gtc = int(r['DonGTC_num'])
        gan_lay = int(r['DonGanLay_num'])
        ltc = int(r['DonLTC_num'])
        l_hh = r['LuongHH_num']
        t_luong = r['TongLuong_num']
        
        # fix display if tong_luong was corrupted in raw string (e.g., 250 vs 250,000)
        actual_luong = l_hh if (t_luong < 1000 and l_hh >= 1000) else t_luong
        
        output_md += f"| {i} | {nv} | {tn} | {nvl} | {gan_giao} | {gtc} | {gan_lay} | {ltc} | {l_hh:,.0f} | {actual_luong:,.0f} |\n"
    
    output_md += "\n"

with open("scratch/income_report.md", "w", encoding="utf-8") as f:
    f.write(output_md)

print("Saved report to scratch/income_report.md")
