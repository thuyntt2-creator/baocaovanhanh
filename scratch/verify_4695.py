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

ws_data = spreadsheet.worksheet("data")
vals_data = ws_data.get_all_values()
df_data = pd.DataFrame(vals_data[1:], columns=vals_data[0])

def clean_num(val):
    if not val: return 0
    try: return float(str(val).replace('.', '').replace(',', '.'))
    except: return 0

for name in ['Di Linh', 'Cam Linh']:
    sub = df_data[(df_data['Chi tiết'].str.contains(name, na=False)) & (df_data['Time'] == '2026-07-22 - Thứ 4')]
    
    tot_vol = sub['Volume'].apply(clean_num).sum()
    tot_gtc = sub['Sản Lượng Giao Thành Công'].apply(clean_num).sum()
    tot_gan = sub['Sản Lượng Gán'].apply(clean_num).sum()
    
    rate_vol = (tot_gtc / tot_vol * 100) if tot_vol > 0 else 0
    rate_gan = (tot_gtc / tot_gan * 100) if tot_gan > 0 else 0
    
    print(f"\n--- {name} (2026-07-22) ---")
    print(f"Tổng Volume (Mới + Tồn): {tot_vol}")
    print(f"Tổng Gán (Mới + Tồn): {tot_gan}")
    print(f"Tổng GTC (Mới + Tồn): {tot_gtc}")
    print(f"-> % GTC / Tổng Volume = {tot_gtc} / {tot_vol} = {rate_vol:.2f}%")
    print(f"-> % GTC / Tổng Gán = {tot_gtc} / {tot_gan} = {rate_gan:.2f}%")

