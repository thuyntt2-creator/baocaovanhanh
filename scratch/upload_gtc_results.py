import os
import io
import sys
import gspread
import pandas as pd
from google.oauth2.service_account import Credentials

# Fix encoding for Windows
os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except AttributeError:
    pass

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PARENT_DIR = os.path.dirname(SCRIPT_DIR)
JSON_FILE = os.path.join(PARENT_DIR, 'credentials.json')
scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

def clean_volume(val):
    if val is None or pd.isna(val) or val == "":
        return 0
    if isinstance(val, (int, float)):
        return int(round(float(val)))
    val_str = str(val).strip()
    if ',' in val_str and '.' in val_str:
        if val_str.index(',') < val_str.index('.'):
            val_clean = val_str.replace(',', '')
        else:
            val_clean = val_str.replace('.', '').replace(',', '.')
    elif ',' in val_str:
        parts = val_str.split(',')
        if len(parts[-1]) == 3:
            val_clean = val_str.replace(',', '')
        else:
            val_clean = val_str.replace(',', '.')
    elif '.' in val_str:
        parts = val_str.split('.')
        if len(parts[-1]) == 3:
            val_clean = val_str.replace('.', '')
        else:
            val_clean = val_str
    else:
        val_clean = val_str
    try:
        return int(round(float(val_clean)))
    except ValueError:
        return 0

def run():
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    
    sheet_key = '1DAwY-46twFrHIs77R4p4IMuIZ6JTE-e58Aj-9Kcr5Jk'
    sh = gc_client.open_by_key(sheet_key)
    ws_data = sh.worksheet("Data")
    data = ws_data.get_all_values()
    df = pd.DataFrame(data[1:], columns=data[0])
    
    # Sort and get unique dates
    all_dates = sorted(df['Time'].unique())
    last_7_dates = all_dates[-7:]
    
    # Filter for the last 7 dates
    df_7d = df[df['Time'].isin(last_7_dates)].copy()
    df_7d['Volume_clean'] = df_7d['Volume'].apply(clean_volume)
    df_7d['GTC_clean'] = df_7d['Sản Lượng Giao Thành Công'].apply(clean_volume)
    
    grouped = df_7d.groupby('Chi tiết').agg({
        'Volume_clean': 'sum',
        'GTC_clean': 'sum',
        'Cấp Quản Lý': 'first',
        'AM': 'first'
    }).reset_index()
    
    grouped['GTC_rate'] = grouped['GTC_clean'] / grouped['Volume_clean']
    
    low_gtc = grouped[grouped['GTC_rate'] < 0.50].sort_values(by='GTC_rate')
    
    # Build list of rows to upload
    upload_data = []
    # Headers
    upload_data.append([
        "STT", "Bưu cục", "Khu vực", "AM Quản lý", "Tỷ lệ GTC (7 ngày)", "Sản lượng GTC", "Tổng sản lượng Volume"
    ])
    
    for idx, (i, row) in enumerate(low_gtc.iterrows()):
        upload_data.append([
            idx + 1,
            row['Chi tiết'],
            row['Cấp Quản Lý'],
            row['AM'],
            f"{row['GTC_rate'] * 100:.2f}%",
            row['GTC_clean'],
            row['Volume_clean']
        ])
        
    tab_name = "GTC 7D dưới 50%"
    try:
        ws_target = sh.worksheet(tab_name)
        print(f"Tab '{tab_name}' already exists. Clearing old data...")
        ws_target.clear()
    except gspread.exceptions.WorksheetNotFound:
        print(f"Tab '{tab_name}' not found. Creating a new one...")
        ws_target = sh.add_worksheet(title=tab_name, rows=100, cols=10)
        
    ws_target.update(upload_data, value_input_option='USER_ENTERED')
    print(f"Successfully uploaded {len(upload_data) - 1} rows to tab '{tab_name}'!")

if __name__ == "__main__":
    run()
