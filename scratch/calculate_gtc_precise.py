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
    ws = sh.worksheet("Data")
    data = ws.get_all_values()
    df = pd.DataFrame(data[1:], columns=data[0])
    
    # Sort and get unique dates
    all_dates = sorted(df['Time'].unique())
    last_7_dates = all_dates[-7:]
    print("Analyzing last 7 dates:")
    for d in last_7_dates:
        print(f"  - {d}")
        
    # Filter for the last 7 dates
    df_7d = df[df['Time'].isin(last_7_dates)].copy()
    
    # Apply clean_volume to both Volume and Sản Lượng Giao Thành Công
    df_7d['Volume_clean'] = df_7d['Volume'].apply(clean_volume)
    df_7d['GTC_clean'] = df_7d['Sản Lượng Giao Thành Công'].apply(clean_volume)
    
    # Sum up Volume and GTC per post office
    grouped = df_7d.groupby('Chi tiết').agg({
        'Volume_clean': 'sum',
        'GTC_clean': 'sum'
    }).reset_index()
    
    # Calculate GTC rate
    grouped['GTC_rate'] = grouped['GTC_clean'] / grouped['Volume_clean']
    
    # Filter for rate < 50%
    low_gtc = grouped[grouped['GTC_rate'] < 0.50].sort_values(by='GTC_rate')
    
    print(f"\nNumber of post offices with GTC < 50% in last 7 days: {len(low_gtc)}")
    for idx, row in low_gtc.iterrows():
        rate_pct = row['GTC_rate'] * 100
        print(f"{row['Chi tiết']}: GTC = {row['GTC_clean']:,} / Vol = {row['Volume_clean']:,} ({rate_pct:.2f}%)")

if __name__ == "__main__":
    run()
