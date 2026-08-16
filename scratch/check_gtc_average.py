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

def clean_pct(val):
    if val is None or pd.isna(val) or val == "":
        return 0.0
    if isinstance(val, (int, float)):
        if val <= 1.0:
            return float(val)
        return float(val) / 100.0
    val_str = str(val).strip().replace('%', '')
    if ',' in val_str and '.' not in val_str:
        val_str = val_str.replace(',', '.')
    try:
        num = float(val_str)
        if num > 1.0:
            return num / 100.0
        return num
    except ValueError:
        return 0.0

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
    
    # Filter for the last 7 dates
    df_7d = df[df['Time'].isin(last_7_dates)].copy()
    
    # Apply cleaning
    df_7d['Volume_clean'] = df_7d['Volume'].apply(clean_volume)
    df_7d['GTC_clean'] = df_7d['Sản Lượng Giao Thành Công'].apply(clean_volume)
    df_7d['pct_gtc_clean'] = df_7d['% GTC'].apply(clean_pct)
    
    # Method 1: Proper Weighted GTC Rate (sum(GTC) / sum(Vol))
    g1 = df_7d.groupby('Chi tiết').agg({
        'Volume_clean': 'sum',
        'GTC_clean': 'sum'
    }).reset_index()
    g1['GTC_rate_weighted'] = g1['GTC_clean'] / g1['Volume_clean']
    
    # Method 2: Simple average of daily % GTC
    # First, let's aggregate % GTC by date and post office (since there are multiple Loại Hàng per day)
    # Wait, how does the sheet calculate % GTC daily for a post office? It would be sum(GTC)/sum(Vol) for that day.
    # Let's group by post office and date, calculate the daily rate, and then average those daily rates.
    df_daily = df_7d.groupby(['Chi tiết', 'Time']).agg({
        'Volume_clean': 'sum',
        'GTC_clean': 'sum'
    }).reset_index()
    df_daily['GTC_rate_daily'] = df_daily['GTC_clean'] / df_daily['Volume_clean']
    
    g2 = df_daily.groupby('Chi tiết')['GTC_rate_daily'].mean().reset_index()
    g2.rename(columns={'GTC_rate_daily': 'GTC_rate_simple_avg'}, inplace=True)
    
    # Merge and compare
    merged = pd.merge(g1, g2, on='Chi tiết')
    
    print("Method 1 (Weighted) < 50%:")
    w_under_50 = merged[merged['GTC_rate_weighted'] < 0.5].sort_values(by='GTC_rate_weighted')
    for idx, r in w_under_50.iterrows():
        print(f"  {r['Chi tiết']}: Weighted={r['GTC_rate_weighted']:.2%}, Simple={r['GTC_rate_simple_avg']:.2%}, Vol={r['Volume_clean']}")
        
    print("\nMethod 2 (Simple Average) < 50%:")
    s_under_50 = merged[merged['GTC_rate_simple_avg'] < 0.5].sort_values(by='GTC_rate_simple_avg')
    for idx, r in s_under_50.iterrows():
        print(f"  {r['Chi tiết']}: Weighted={r['GTC_rate_weighted']:.2%}, Simple={r['GTC_rate_simple_avg']:.2%}, Vol={r['Volume_clean']}")

if __name__ == "__main__":
    run()
