import sys
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials

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

# Support employees
ws_hotro = spreadsheet.worksheet("NV HỖ TRỢ")
vals_hotro = ws_hotro.get_all_values()
support_emp_codes = set()
support_emp_names = set()
for r in vals_hotro:
    if r and r[0].strip():
        raw_val = r[0].strip()
        code = raw_val.split('_')[0].split('-')[0].strip()
        if code.isdigit():
            support_emp_codes.add(code)
        name = raw_val.split('_')[-1].split('-')[-1].strip()
        if name and not name.isdigit():
            support_emp_names.add(name.lower())

# Năng suất
ns_vals = spreadsheet.worksheet("năng suất").get_all_values()
df_ns = pd.DataFrame(ns_vals[1:], columns=ns_vals[0])
df_ns['Emp_Code'] = df_ns['NhanVien'].apply(lambda x: str(x).split('_')[0].split('-')[0].strip())
df_ns['Emp_Name'] = df_ns['NhanVien'].apply(lambda x: str(x).split('_')[-1].split('-')[-1].strip())

for hub_code, hub_name in [("22830000", "Cam Linh"), ("20942000", "Di Linh")]:
    sub_ns_hub = df_ns[
        df_ns['Bưu cục'].str.contains(hub_code, regex=False, na=False) |
        df_ns['Bưu cục'].str.contains(hub_name, regex=False, na=False)
    ].copy()
    if len(sub_ns_hub) > 0:
        latest_date = sub_ns_hub['Ngay'].unique()[0]
        sub_ns = sub_ns_hub[sub_ns_hub['Ngay'] == latest_date]
        total_ns_all = len(sub_ns)
        
        sup_ns = sub_ns[
            (sub_ns['Emp_Code'].isin(support_emp_codes)) |
            (sub_ns['Emp_Name'].str.lower().isin(support_emp_names))
        ]
        sup_cnt = len(sup_ns)
        reg_cnt = total_ns_all - sup_cnt
        
        print(f"=== {hub_name} ({hub_code}) on {latest_date} ===")
        print(f"Total employees in 'năng suất': {total_ns_all} NV ({reg_cnt} PTTT + {sup_cnt} NV hỗ trợ)")
