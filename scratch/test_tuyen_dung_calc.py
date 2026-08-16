import sys
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials

sys.stdout.reconfigure(encoding='utf-8')

SHEET_ID = "1MtbZBgRFwCWj6uQKsSqddiJ2GsTiEvKxRIPSshDa5PM"
CREDENTIALS_FILE = "credentials.json"

scope = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]
creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
client = gspread.authorize(creds)
spreadsheet = client.open_by_key(SHEET_ID)

rec_vals = spreadsheet.worksheet("báo cáo tuyển dụng").get_all_values()
df_rec = pd.DataFrame(rec_vals)

for hub_code, hub_name in [("22830000", "Cam Linh"), ("20942000", "Di Linh")]:
    rec_row = None
    for idx, r in df_rec.iterrows():
        row_str = " ".join([str(x) for x in r])
        if hub_code in row_str or hub_name.lower() in row_str.lower():
            rec_row = r
            break
    if rec_row is not None:
        stt_pttt = rec_row[7]
        thieu_pttt = float(rec_row[8]) if rec_row[8].isdigit() else 0
        hientai_pttt = float(rec_row[9]) if rec_row[9].isdigit() else 0
        dinhbien_pttt = float(rec_row[10]) if rec_row[10].isdigit() else 0
        
        calc_hientai = dinhbien_pttt - thieu_pttt if dinhbien_pttt > 0 else hientai_pttt
        
        print(f"=== {hub_name} ({hub_code}) ===")
        print(f"Sheet tuyển dụng: Định biên = {dinhbien_pttt}, Hiện tại (cell) = {hientai_pttt}, Thiếu = {thieu_pttt}")
        print(f"Tính chuẩn: Hiện tại = Định biên ({dinhbien_pttt}) - Thiếu ({thieu_pttt}) = {calc_hientai}")
