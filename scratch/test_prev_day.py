import os
import sys
import pandas as pd
import gspread
from datetime import datetime, timedelta
from google.oauth2.service_account import Credentials

try:
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    if hasattr(sys.stderr, 'reconfigure'):
        sys.stderr.reconfigure(encoding='utf-8')
except Exception:
    pass

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(BASE_DIR)

MAIN_SPREADSHEET_ID = "1JZ1eRerRqrpwjZ4HBevQunjd8VquM_cvPFz12TaJfMQ"
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def main():
    credentials = Credentials.from_service_account_file(JSON_FILE, scopes=SCOPES)
    gc = gspread.authorize(credentials)
    sh = gc.open_by_key(MAIN_SPREADSHEET_ID)
    
    # We will test for 2026-07-02
    target_date = datetime(2026, 7, 2)
    prev_date = target_date - timedelta(days=1)
    
    weekday_map = {
        0: "Thứ 2", 1: "Thứ 3", 2: "Thứ 4", 3: "Thứ 5", 4: "Thứ 6", 5: "Thứ 7", 6: "Chủ Nhật"
    }
    date_str_sheet = f"{target_date.strftime('%Y-%m-%d')} - {weekday_map[target_date.weekday()]}"
    prev_date_str_sheet = f"{prev_date.strftime('%Y-%m-%d')} - {weekday_map[prev_date.weekday()]}"
    
    print(f"Target date (N-1): {date_str_sheet}")
    print(f"Prev date (N-2): {prev_date_str_sheet}")
    
    # Load CoCauVung
    ws_cc = sh.worksheet("CoCauVung")
    cc_rows = ws_cc.get_all_values()
    df_cc = pd.DataFrame(cc_rows[1:], columns=cc_rows[0])
    df_cc = df_cc[df_cc['Bưu cục'].str.strip() != '']
    master_pos = df_cc['Bưu cục'].str.strip().tolist()
    
    # Load Data
    print("Loading Data...")
    ws_data = sh.worksheet("Data")
    data_rows = ws_data.get_all_values()
    df_data = pd.DataFrame(data_rows[1:], columns=data_rows[0])
    
    df_data_n1 = df_data[df_data['Time'] == date_str_sheet]
    df_data_n2 = df_data[df_data['Time'] == prev_date_str_sheet]
    print(f"Loaded {len(df_data_n1)} rows for N-1, {len(df_data_n2)} rows for N-2")
    
    # Helper to calculate GTC rate per PO
    def get_gtc_rates(df):
        po_gtc = {}
        for idx, row in df.iterrows():
            raw_po = row['Chi tiết']
            # We match to master_pos
            from generate_morning_questions import match_po_name, parse_int
            std_po = match_po_name(raw_po, master_pos)
            if not std_po:
                continue
            if std_po not in po_gtc:
                po_gtc[std_po] = {'vol': 0, 'gtc': 0}
            po_gtc[std_po]['vol'] += parse_int(row['Volume'])
            po_gtc[std_po]['gtc'] += parse_int(row['Sản Lượng Giao Thành Công'])
        
        rates = {}
        for po, info in po_gtc.items():
            if info['vol'] > 0:
                rates[po] = (info['gtc'] / info['vol']) * 100.0
        return rates

    rates_n1 = get_gtc_rates(df_data_n1)
    rates_n2 = get_gtc_rates(df_data_n2)
    
    print("\n--- Day-over-day GTC changes (Top drops) ---")
    drops = []
    for po in master_pos:
        r1 = rates_n1.get(po)
        r2 = rates_n2.get(po)
        if r1 is not None and r2 is not None:
            diff = r1 - r2
            if diff < 0: # Drop
                drops.append((po, r2, r1, -diff))
    
    drops.sort(key=lambda x: -x[3])
    for item in drops[:10]:
        print(f"PO: {item[0]:<30} Yesterday: {item[1]:.2f}% Today: {item[2]:.2f}% Drop: {item[3]:.2f}%")

if __name__ == "__main__":
    main()
