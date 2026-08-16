import sys
import io
import os
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import json
import unicodedata

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
SHEET_KEY = '14r8n9L2cIG1Bmz8kSH79B24QzmnOApZhniGyOU40hr4'
RAW_EXCEL = r"C:\Users\lap4all\.gemini\antigravity-ide\scratch\Rớt LC 22_6_2026 - Full sàn.xlsx"
scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def normalize_str(s):
    if not s:
        return ""
    return unicodedata.normalize('NFC', str(s).strip())

def clean_bc_name(name):
    name = normalize_str(name).lower()
    for tag in ['(dno)', '(ldo)', '(kho)', '(bth)', '(nth)']:
        name = name.replace(tag, "")
    for prefix in ['kho chuyển tiếp', 'kho trung chuyển', 'điểm xử lý hàng', 'điểm lấy hàng', 'bưu cục', 'bc', 'đl']:
        name = name.replace(prefix, "")
    return name.strip()

def main():
    credentials = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(credentials)
    sh = gc_client.open_by_key(SHEET_KEY)
    
    ws_data = sh.worksheet("Data")
    data_rows = ws_data.get_all_values()
    df_data = pd.DataFrame(data_rows[1:], columns=data_rows[0])
    
    df_22 = df_data[df_data['Loại ngày'] == '2026-06-22'].copy()
    
    # Load raw Excel
    df_excel = pd.read_excel(RAW_EXCEL, sheet_name="Sheet1")
    df_excel_ntb = df_excel[(df_excel['vung_xuat'] == 'NTB') & (df_excel['loai_kh'] == 'TTS')]
    
    print(f"Raw NTB TTS rows: {len(df_excel_ntb)}")
    
    # Map raw Excel rows to Data sheet PO names
    # Let's see: we want to find for each row in raw Excel which Data row it matches
    po_names_22 = df_22['Chi tiết'].tolist()
    po_names_22_clean = [clean_bc_name(x) for x in po_names_22]
    
    match_map = {}
    unmatched_count = 0
    
    for idx, row in df_excel_ntb.iterrows():
        raw_po = row['tenbcxuat']
        raw_po_clean = clean_bc_name(raw_po)
        
        # Try direct match
        found = False
        for name in po_names_22:
            if normalize_str(raw_po).lower() == normalize_str(name).lower():
                match_map[name] = match_map.get(name, 0) + 1
                found = True
                break
        if found:
            continue
            
        # Try clean match
        for name, clean_name in zip(po_names_22, po_names_22_clean):
            if raw_po_clean and clean_name and (raw_po_clean in clean_name or clean_name in raw_po_clean):
                match_map[name] = match_map.get(name, 0) + 1
                found = True
                break
                
        if not found:
            print(f"Could not map raw Excel PO: '{raw_po}'")
            unmatched_count += 1
            
    print(f"Successfully matched: {sum(match_map.values())} rows")
    print(f"Unmatched count: {unmatched_count}")
    
    print("\nMatched counts per PO:")
    for name, count in sorted(match_map.items()):
        print(f"  '{name}': {count}")

if __name__ == "__main__":
    main()
