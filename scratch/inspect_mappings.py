import sys
import io
import os
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import unicodedata

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
SHEET_KEY = '14r8n9L2cIG1Bmz8kSH79B24QzmnOApZhniGyOU40hr4'
scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def normalize_str(s):
    if not s:
        return ""
    return unicodedata.normalize('NFC', str(s).strip())

def main():
    credentials = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(credentials)
    sh = gc_client.open_by_key(SHEET_KEY)
    
    # Read Data sheet
    ws_data = sh.worksheet("Data")
    data_rows = ws_data.get_all_values()
    df_data = pd.DataFrame(data_rows[1:], columns=data_rows[0])
    
    # Read Cơ cấu sheet
    ws_cocau = sh.worksheet("Cơ cấu")
    cocau_rows = ws_cocau.get_all_values()
    df_cocau = pd.DataFrame(cocau_rows[1:], columns=cocau_rows[0])
    
    print(f"Data rows: {len(df_data)}")
    print(f"Cơ cấu rows: {len(df_cocau)}")
    
    unique_data_po = df_data['Chi tiết'].unique()
    print(f"Unique POs in Data: {len(unique_data_po)}")
    
    # Let's inspect mapping by exact name and by normalized/cleaned substrings
    cocau_bcs = df_cocau['Bưu cục'].tolist()
    cocau_bcs_norm = [normalize_str(x) for x in cocau_bcs]
    
    matched = 0
    unmatched_list = []
    
    for po in unique_data_po:
        po_norm = normalize_str(po)
        # Try direct match
        found = False
        for idx, bc in enumerate(cocau_bcs_norm):
            if po_norm == bc:
                matched += 1
                found = True
                break
        if found:
            continue
            
        # Try substring match: check if po is substring of bc, or bc is substring of po
        # E.g. cleaning words like "Bưu Cục", "Điểm Xử Lý", "BC"
        clean_po = po_norm.replace("Bưu Cục", "").replace("Bưu cục", "").replace("BC", "").replace("Điểm Xử Lý", "").replace("Điểm lấy hàng", "").strip()
        
        for idx, bc in enumerate(cocau_bcs_norm):
            clean_bc = bc.replace("Bưu Cục", "").replace("Bưu cục", "").replace("BC", "").replace("Điểm Xử Lý", "").replace("Điểm lấy hàng", "").strip()
            if clean_po and clean_bc and (clean_po in clean_bc or clean_bc in clean_po):
                matched += 1
                found = True
                break
        if not found:
            unmatched_list.append(po)
            
    print(f"Matched: {matched} / {len(unique_data_po)}")
    print(f"Unmatched list ({len(unmatched_list)}):")
    for u in unmatched_list:
        print(f"- '{u}'")

if __name__ == "__main__":
    main()
