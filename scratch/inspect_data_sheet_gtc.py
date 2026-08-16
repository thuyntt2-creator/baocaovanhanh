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

data_vals = spreadsheet.worksheet("data").get_all_values()
df_data = pd.DataFrame(data_vals[1:], columns=data_vals[0])

for hub_code, hub_name in [("22830000", "Cam Linh"), ("20942000", "Di Linh")]:
    print(f"\n=================== {hub_name} ({hub_code}) ===================")
    sub_d = df_data[
        df_data['Chi tiết'].str.contains(hub_code, regex=False, na=False) |
        df_data['Chi tiết'].str.contains(hub_name, regex=False, na=False)
    ]
    for d in sorted(sub_d['Time'].unique(), reverse=True)[:10]:
        sub_d_date = sub_d[sub_d['Time'] == d]
        tot_vol = sum(clean_num(r['Volume']) for _, r in sub_d_date.iterrows())
        tot_gtc = sum(clean_num(r['Sản Lượng Giao Thành Công']) for _, r in sub_d_date.iterrows())
        tot_gan = sum(clean_num(r['Sản Lượng Gán']) for _, r in sub_d_date.iterrows())
        rate_vol = (tot_gtc / tot_vol * 100) if tot_vol > 0 else 0
        rate_gan = (tot_gtc / tot_gan * 100) if tot_gan > 0 else 0
        print(f"Date '{d}': Volume={tot_vol}, Gán={tot_gan}, GTC={tot_gtc} | GTC/Volume = {rate_vol:.2f}%, GTC/Gán = {rate_gan:.2f}%")
