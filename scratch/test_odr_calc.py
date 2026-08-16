import os
import sys
import gspread
import pandas as pd
from google.oauth2.service_account import Credentials

sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = r"c:\Users\lap4all\Documents\Auto report"
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
SHEET_KEY = '1705_0rKkgXBpsCbgK10EDr_mzSGhJOAcCa1WZsrWrU4'

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def clean_bc_name(name):
    if not name:
        return ""
    name = str(name).strip().lower()
    for tag in ['(dno)', '(ldo)', '(kho)', '(bth)', '(nth)']:
        name = name.replace(tag, "")
    for prefix in ['kho chuyển tiếp', 'kho trung chuyển', 'điểm xử lý hàng', 'điểm lấy hàng', 'bưu cục', 'bc', 'đl']:
        name = name.replace(prefix, "")
    return name.strip()

def resolve_po_info(po_name, cocau_map):
    po_key = str(po_name).strip().lower()
    
    # 1. Direct match
    if po_key in cocau_map:
        return cocau_map[po_key]
        
    # 2. Match after cleaning
    clean_key = clean_bc_name(po_name)
    for k, v in cocau_map.items():
        if clean_bc_name(k) == clean_key:
            return v
            
    # 3. Fuzzy match (substring)
    for k, v in cocau_map.items():
        clean_k = clean_bc_name(k)
        if clean_key and clean_k and (clean_key in clean_k or clean_k in clean_key):
            return v
            
    # Fallbacks based on typical PO prefixes if available
    if "dno" in po_key or "đắk nông" in po_key:
        return ("Trần Văn Phước", "Đắk Nông")  # Default AM for Đắk Nông
    return ("Chưa gán AM", "Chưa gán Tỉnh")

def main():
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SHEET_KEY)
    
    # --- 1. Load Sheets ---
    print("Loading worksheets...")
    df_odr = pd.DataFrame(sh.worksheet("ODR").get_all_records())
    df_cocau = pd.DataFrame(sh.worksheet("Cocau").get_all_records())
    df_tts = pd.DataFrame(sh.worksheet("ODR - TTS").get_all_records())
    
    # --- 2. Build Cơ cấu map ---
    cocau_map = {}
    for idx, row in df_cocau.iterrows():
        bc_name = str(row['Bưu cục']).strip().lower()
        cocau_map[bc_name] = (row['AM'], row['Tỉnh'])
        
    # --- 3. Calculate ODR TTS by AM ---
    # target_date in ODR sheet: let's pick the latest date in the sheet
    # Format of Time column is e.g. "2026-06-27 - Thứ 7"
    df_odr['date_parsed'] = df_odr['Time'].apply(lambda x: x.split(" - ")[0].strip() if " - " in str(x) else str(x).strip())
    latest_date_str = df_odr['date_parsed'].max()
    print(f"\nLatest date found in ODR sheet: {latest_date_str}")
    
    # Filter for latest date
    df_odr_latest = df_odr[df_odr['date_parsed'] == latest_date_str].copy()
    print(f"Number of rows for latest date: {len(df_odr_latest)}")
    
    # Map AM to each row
    ams = []
    provinces = []
    for idx, row in df_odr_latest.iterrows():
        po_name = row['Chi tiết']
        am, prov = resolve_po_info(po_name, cocau_map)
        ams.append(am)
        provinces.append(prov)
    df_odr_latest['AM'] = ams
    df_odr_latest['Tỉnh'] = provinces
    
    # Parse GTC and %Ontime
    df_odr_latest['GTC'] = pd.to_numeric(df_odr_latest['GTC'], errors='coerce').fillna(0)
    def parse_percent(val):
        if not val:
            return 1.0
        val_str = str(val).replace(",", ".").replace("%", "").strip()
        try:
            return float(val_str) / 100
        except ValueError:
            return 1.0
            
    df_odr_latest['Ontime_rate'] = df_odr_latest['%Ontime'].apply(parse_percent)
    df_odr_latest['Ontime_volume'] = df_odr_latest['GTC'] * df_odr_latest['Ontime_rate']
    
    # Group by AM to calculate ODR rate
    am_odr = df_odr_latest.groupby('AM').agg(
        total_gtc=('GTC', 'sum'),
        ontime_gtc=('Ontime_volume', 'sum')
    ).reset_index()
    am_odr['ODR_Rate'] = (am_odr['ontime_gtc'] / am_odr['total_gtc'] * 100).fillna(100.0)
    am_odr = am_odr.sort_values(by='ODR_Rate', ascending=False).reset_index(drop=True)
    
    print("\n--- TABLE 1: ODR TTS BY AM ---")
    print(am_odr.to_string())
    
    # --- 4. Process ODR - TTS (Backlog) ---
    print("\nProcessing backlog (ODR - TTS)...")
    df_tts['total_orders'] = pd.to_numeric(df_tts['total_orders'], errors='coerce').fillna(0).astype(int)
    
    # Pivot table to group by AM and nhom_tre
    backlog_pivot = df_tts.pivot_table(
        index='AM',
        columns='nhom_tre',
        values='total_orders',
        aggfunc='sum',
        fill_value=0
    ).reset_index()
    
    # Add Grand Total column
    delay_cols = [c for c in backlog_pivot.columns if c != 'AM']
    backlog_pivot['Tổng tồn'] = backlog_pivot[delay_cols].sum(axis=1)
    backlog_pivot = backlog_pivot.sort_values(by='Tổng tồn', ascending=False).reset_index(drop=True)
    
    print("\n--- TABLE 2: BACKLOG ODR TTS BY AM ---")
    print(backlog_pivot.to_string())

if __name__ == "__main__":
    main()
