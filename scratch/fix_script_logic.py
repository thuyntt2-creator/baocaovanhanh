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
    s = str(val).strip()
    # If standard Vietnamese number format: "224.458,50" -> 224458.5
    # If "1,122" or "1.122" thousand separator -> 1122
    # Remove dots if followed by 3 digits
    s = s.replace('đ', '').replace('%', '').strip()
    if ',' in s and '.' in s:
        # e.g. 224.458,50 or 1,122.50
        if s.rfind(',') > s.rfind('.'): # 224.458,50
            s = s.replace('.', '').replace(',', '.')
        else: # 1,122.50
            s = s.replace(',', '')
    elif ',' in s: # e.g. 59,54 or 1,122
        # Check if 3 digits after comma -> thousand separator e.g. 1,122
        parts = s.split(',')
        if len(parts) == 2 and len(parts[1]) == 3 and not parts[1].endswith('00'):
            s = s.replace(',', '')
        else:
            s = s.replace(',', '.')
    elif '.' in s: # e.g. 224.458 or 59.54
        parts = s.split('.')
        if len(parts) > 1 and all(len(p) == 3 for p in parts[1:]):
            s = s.replace('.', '')
    try:
        return float(s)
    except ValueError:
        return 0.0

print("Test clean_num:")
print("224.458 =>", clean_num("224.458"))
print("478.443 =>", clean_num("478.443"))
print("1,122 =>", clean_num("1,122"))
print("59,54% =>", clean_num("59,54%"))

ws_data = spreadsheet.worksheet("data")
df_data = pd.DataFrame(ws_data.get_all_values()[1:], columns=ws_data.get_all_values()[0])

ws_tn = spreadsheet.worksheet("thu nhập")
df_tn = pd.DataFrame(ws_tn.get_all_values()[1:], columns=ws_tn.get_all_values()[0])

ws_ns = spreadsheet.worksheet("năng suất")
df_ns = pd.DataFrame(ws_ns.get_all_values()[1:], columns=ws_ns.get_all_values()[0])

for query in ['20942000', 'Di Linh', '22830000', 'Cam Linh']:
    print(f"\n==================== SEARCH FOR '{query}' ====================")
    # Search Data
    sub_d = df_data[df_data.apply(lambda r: r.astype(str).str.contains(query, case=False).any(), axis=1)]
    print(f"Data sheet matches: {len(sub_d)} rows")
    if len(sub_d) > 0:
        latest_date = sub_d['Time'].max()
        print(f"  Latest date in Data: {latest_date}")
        sub_latest = sub_d[sub_d['Time'] == latest_date]
        for _, r in sub_latest.iterrows():
            print(f"  Loại hàng: {r['Loại Hàng']} | Vol: {r['Volume']} | GTC: {r['Sản Lượng Giao Thành Công']}")
            
    # Search Thu Nhập
    sub_tn = df_tn[df_tn.apply(lambda r: r.astype(str).str.contains(query, case=False).any(), axis=1)]
    print(f"Thu nhập sheet matches: {len(sub_tn)} rows")
    
    # Search Năng Suất
    sub_ns = df_ns[df_ns.apply(lambda r: r.astype(str).str.contains(query, case=False).any(), axis=1)]
    print(f"Năng suất sheet matches: {len(sub_ns)} rows")
    if len(sub_ns) > 0:
        latest_ns_date = sub_ns['Ngay'].unique()
        print(f"  Dates in Năng suất: {latest_ns_date[:5]}")

