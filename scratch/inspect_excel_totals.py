import sys
import io
import os
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
SHEET_KEY = '14r8n9L2cIG1Bmz8kSH79B24QzmnOApZhniGyOU40hr4'
scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def main():
    credentials = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(credentials)
    sh = gc_client.open_by_key(SHEET_KEY)
    
    ws_data = sh.worksheet("Data")
    data_rows = ws_data.get_all_values()
    df_data = pd.DataFrame(data_rows[1:], columns=data_rows[0])
    
    df_data['Vol cần LC'] = pd.to_numeric(df_data['Vol cần LC'], errors='coerce').fillna(0)
    df_data['%_rot_lc'] = pd.to_numeric(df_data['%_rot_lc'], errors='coerce').fillna(0)
    # Vol rớt LC = Vol cần LC * %_rot_lc
    df_data['Vol rớt LC'] = df_data['Vol cần LC'] * df_data['%_rot_lc']
    
    for date in ['2026-06-22', '2026-06-21']:
        df_date = df_data[df_data['Loại ngày'] == date]
        tot_need = df_date['Vol cần LC'].sum()
        tot_drop = df_date['Vol rớt LC'].sum()
        rate = (tot_drop / tot_need * 100) if tot_need > 0 else 0.0
        print(f"Date: {date}")
        print(f"  Total Vol cần LC: {tot_need:,.0f}")
        print(f"  Total Vol rớt LC: {tot_drop:,.0f}")
        print(f"  Rate: {rate:.4f}%")

if __name__ == "__main__":
    main()
