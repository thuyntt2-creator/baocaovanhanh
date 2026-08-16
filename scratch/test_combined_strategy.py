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
    
    ws_cocau = sh.worksheet("Cơ cấu")
    cocau_rows = ws_cocau.get_all_values()
    df_cocau = pd.DataFrame(cocau_rows[1:], columns=cocau_rows[0])
    
    unique_data_po = df_data['Chi tiết'].unique()
    
    # 1. Build Cơ cấu lookup
    cocau_map = {}
    for idx, row in df_cocau.iterrows():
        bc_name = normalize_str(row['Bưu cục'])
        cocau_map[bc_name.lower()] = (row['AM'], row['Tỉnh'])
        
    # 2. Load JSON mapping
    json_path = os.path.join(BASE_DIR, "scratch", "extracted_mappings.json")
    with open(json_path, "r", encoding="utf-8") as f:
        known_mappings = json.load(f)
    std_mappings = {normalize_str(k).lower(): v for k, v in known_mappings.items()}
    
    matched = 0
    unmatched = []
    
    for po in unique_data_po:
        po_norm = normalize_str(po)
        po_key = po_norm.lower()
        
        # Strategy A: Direct match in Cơ cấu
        if po_key in cocau_map:
            am, prov = cocau_map[po_key]
            matched += 1
            print(f"Direct Cơ cấu match: '{po}' -> AM: {am}, Province: {prov}")
            continue
            
        # Strategy B: Direct match in JSON mappings
        if po_key in std_mappings:
            am, prov = std_mappings[po_key]
            matched += 1
            print(f"Direct JSON match: '{po}' -> AM: {am}, Province: {prov}")
            continue
            
        # Strategy C: Fuzzy match in JSON mappings (substring check)
        found_json = None
        clean_po = po_key.replace("bưu cục", "").replace("bc", "").replace("bưu cục", "").replace(" ", "")
        for k, v in std_mappings.items():
            clean_k = k.replace("bưu cục", "").replace("bc", "").replace("bưu cục", "").replace(" ", "")
            if clean_po == clean_k or clean_po in clean_k or clean_k in clean_po:
                found_json = v
                break
                
        if found_json:
            matched += 1
            print(f"Fuzzy JSON match: '{po}' -> AM: {found_json[0]}, Province: {found_json[1]}")
            continue
            
        # Strategy D: Fuzzy clean match in Cơ cấu (excluding Kho/Diên Khánh issues)
        # Check if any clean_bc name is in po_key
        found_cocau = []
        for bc_name, info in cocau_map.items():
            clean_bc = clean_bc_name(bc_name)
            if clean_bc and clean_bc in po_key:
                found_cocau.append((clean_bc, info))
                
        if found_cocau:
            # Sort by clean_bc length descending (longest substring is most specific)
            found_cocau.sort(key=lambda x: len(x[0]), reverse=True)
            matched += 1
            print(f"Fuzzy Cơ cấu match (longest clean): '{po}' -> AM: {found_cocau[0][1][0]}, Province: {found_cocau[0][1][1]}")
            continue
            
        unmatched.append(po)
        
    print(f"\nMatched {matched} / {len(unique_data_po)}")
    print(f"Unmatched count: {len(unmatched)}")
    for u in unmatched:
        print(f"  - '{u}'")

if __name__ == "__main__":
    main()
