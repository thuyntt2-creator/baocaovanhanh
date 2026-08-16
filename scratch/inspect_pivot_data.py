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
    
    df_22 = df_data[df_data['Loại ngày'] == '2026-06-22']
    print(f"Number of rows for 2026-06-22: {len(df_22)}")
    print(df_22.head(10))
    
    df_21 = df_data[df_data['Loại ngày'] == '2026-06-21']
    print(f"Number of rows for 2026-06-21: {len(df_21)}")
    print(df_21.head(10))

if __name__ == "__main__":
    main()
