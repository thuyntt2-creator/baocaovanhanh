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

if not target_hubs:
    target_hubs = ["(KHO) Cam Linh", "(LDO) Đơn Dương"]

tn_vals = spreadsheet.worksheet("thu nhập").get_all_values()
df_tn = pd.DataFrame(tn_vals[1:], columns=tn_vals[0])

df_tn['LuongHH_num'] = df_tn['Lương HH/ ngày'].apply(clean_num)
df_tn['DonGTC_num'] = df_tn['Đơn giao tính lương'].apply(clean_num)
df_tn['DonLTC_num'] = df_tn['Đơn lấy tính lương'].apply(clean_num)

rec_vals = spreadsheet.worksheet("báo cáo tuyển dụng").get_all_values()
df_rec = pd.DataFrame(rec_vals)

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
    
    date_val = sub['Thời gian'].iloc[0] if len(sub) > 0 else "25-07-2026"
    print(f"\n==========================================")
    print(f"BƯU CỤC: {full_bc_name} - Thu nhập ngày {date_val}")
    print(f"==========================================")
    
    for i, (_, r) in enumerate(sub.iterrows(), 1):
        nv = r['Nhân viên']
        gtc = int(r['DonGTC_num'])
        ltc = int(r['DonLTC_num'])
        l_hh = r['LuongHH_num']
        tn = r['Thâm niên']
        print(f"{i:2d}. {nv:<32} | GTC: {gtc:3d} | LTC: {ltc:2d} | Thu nhập: {l_hh:>9,.} đ | {tn}")
