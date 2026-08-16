import os
import sys
import io
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except AttributeError:
    pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PARENT_DIR = os.path.dirname(BASE_DIR)
JSON_FILE = os.path.join(PARENT_DIR, 'credentials.json')
SHEET_KEY = '1JZ1eRerRqrpwjZ4HBevQunjd8VquM_cvPFz12TaJfMQ'

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def main():
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    sh = gc_client.open_by_key(SHEET_KEY)
    
    # 1. Sheet trên10kg
    ws1 = sh.worksheet('trên10kg')
    data1 = ws1.get_all_records()
    df1 = pd.DataFrame(data1)
    print("=== SHEET: trên10kg ===")
    print("Columns:", list(df1.columns))
    print("Shape:", df1.shape)
    print("Dates min-max:", df1['hen_lay'].min(), "to", df1['hen_lay'].max())
    print("Unique nhom_kh:", df1['nhom_kh'].unique())
    print("Unique nhom_kl:", df1['nhom_kl'].unique())
    print("Unique province_name:", df1['province_name'].unique())
    print("Total warehouses (warehouse_name):", df1['warehouse_name'].nunique())
    print("Sample sum so_don by hen_lay (last 5 dates):")
    df1['so_don_num'] = pd.to_numeric(df1['so_don'], errors='coerce').fillna(0)
    print(df1.groupby('hen_lay')['so_don_num'].sum().tail(10))

    # 2. Sheet SL > 10kg
    ws2 = sh.worksheet('SL > 10kg')
    data2 = ws2.get_all_records()
    df2 = pd.DataFrame(data2)
    print("\n=== SHEET: SL > 10kg ===")
    print("Columns:", list(df2.columns))
    print("Shape:", df2.shape)
    print("Unique Time/Date sample:", df2['Time'].unique()[:10])
    print("Unique Loại Khối Lượng:", df2['Loại Khối Lượng'].unique())
    print("Unique Loại Hàng:", df2['Loại Hàng'].unique())
    print("Unique Bưu cục (Chi tiết):", df2['Chi tiết'].nunique() if 'Chi tiết' in df2.columns else 'N/A')
    if 'Volume' in df2.columns:
        df2['Volume_num'] = pd.to_numeric(df2['Volume'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
        print("Sample Volume sum by Time (last 10):")
        print(df2.groupby('Time')['Volume_num'].sum().tail(10))

if __name__ == "__main__":
    main()
