import sys
import os
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

# 1. Load NV HỖ TRỢ
ws_hotro = None
for ws in spreadsheet.worksheets():
    if any(k in ws.title.lower() for k in ["hỗ trợ", "ho tro", "ho_tro", "hỗ_trợ"]):
        ws_hotro = ws
        break

support_emp_codes = set()
support_emp_names = set()

if ws_hotro:
    vals_hotro = ws_hotro.get_all_values()
    for row in vals_hotro:
        if row and row[0].strip():
            raw_val = row[0].strip()
            if raw_val.lower() in ['nhân viên', 'nhan vien', 'stt', 'mã nv']:
                continue
            # Extract code and name
            # Format usually: 3139329_Đỗ Thanh Khiêm or 3094287-Nguyễn Đức Quyền
            code = raw_val.split('_')[0].split('-')[0].strip()
            if code.isdigit():
                support_emp_codes.add(code)
            name = raw_val.split('_')[-1].split('-')[-1].strip()
            if name and not name.isdigit():
                support_emp_names.add(name.lower())

print(f"Found {len(support_emp_codes)} support employee codes: {sorted(list(support_emp_codes))}")
print(f"Found {len(support_emp_names)} support employee names: {sorted(list(support_emp_names))}")

# 2. Check Data from thu nhập and năng suất
tn_vals = spreadsheet.worksheet("thu nhập").get_all_values()
df_tn = pd.DataFrame(tn_vals[1:], columns=tn_vals[0])
df_tn['Emp_Code'] = df_tn['Nhân viên'].apply(lambda x: str(x).split('-')[0].split('_')[0].strip())
df_tn['Emp_Name'] = df_tn['Nhân viên'].apply(lambda x: str(x).split('-')[-1].split('_')[-1].strip())
df_tn['luong_num'] = df_tn['Lương HH/ ngày'].apply(clean_num)

ns_vals = spreadsheet.worksheet("năng suất").get_all_values()
df_ns = pd.DataFrame(ns_vals[1:], columns=ns_vals[0])
df_ns['Emp_Code'] = df_ns['NhanVien'].apply(lambda x: str(x).split('_')[0].split('-')[0].strip())
df_ns['TongDonGTC_num'] = df_ns['TongDonGTC'].apply(clean_num)

for hub_code, hub_name in [("22830000", "Cam Linh"), ("20942000", "Di Linh")]:
    print(f"\n--- HUB: {hub_name} ({hub_code}) ---")
    sub_tn = df_tn[df_tn['Bưu cục'].str.contains(hub_code, regex=False, na=False) | df_tn['Bưu cục'].str.contains(hub_name, regex=False, na=False)]
    print(f"Total employees in 'thu nhập' before filter: {len(sub_tn)}")
    
    # Filter out support employees
    sub_tn_filtered = sub_tn[~sub_tn['Emp_Code'].isin(support_emp_codes)]
    filtered_out_tn = sub_tn[sub_tn['Emp_Code'].isin(support_emp_codes)]
    print(f"Total employees in 'thu nhập' after filter: {len(sub_tn_filtered)}")
    if len(filtered_out_tn) > 0:
        print("  Excluded support employees in 'thu nhập':")
        for _, r in filtered_out_tn.iterrows():
            print(f"   - {r['Nhân viên']} (Lương: {r['Lương HH/ ngày']})")

    sub_ns_hub = df_ns[df_ns['Bưu cục'].str.contains(hub_code, regex=False, na=False) | df_ns['Bưu cục'].str.contains(hub_name, regex=False, na=False)]
    if len(sub_ns_hub) > 0:
        latest_date = sub_ns_hub['Ngay'].unique()[0] # or sort date
        sub_ns = sub_ns_hub[sub_ns_hub['Ngay'] == latest_date]
        print(f"Total employees in 'năng suất' ({latest_date}) before filter: {len(sub_ns)}")
        sub_ns_filtered = sub_ns[~sub_ns['Emp_Code'].isin(support_emp_codes)]
        filtered_out_ns = sub_ns[sub_ns['Emp_Code'].isin(support_emp_codes)]
        print(f"Total employees in 'năng suất' after filter: {len(sub_ns_filtered)}")
        if len(filtered_out_ns) > 0:
            print("  Excluded support employees in 'năng suất':")
            for _, r in filtered_out_ns.iterrows():
                print(f"   - {r['NhanVien']} (GTC: {r['TongDonGTC']})")
