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

def main():
    credentials = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(credentials)
    sh = gc_client.open_by_key(SHEET_KEY)
    
    ws_data = sh.worksheet("Data")
    data_rows = ws_data.get_all_values()
    df_data = pd.DataFrame(data_rows[1:], columns=data_rows[0])
    
    unique_data_po = df_data['Chi tiết'].unique()
    
    # Load JSON mapping
    json_path = os.path.join(BASE_DIR, "scratch", "extracted_mappings.json")
    with open(json_path, "r", encoding="utf-8") as f:
        known_mappings = json.load(f)
        
    # Standardize keys in known_mappings to normalized NFC lowercase
    std_mappings = {normalize_str(k).lower(): v for k, v in known_mappings.items()}
    
    matched = 0
    unmatched = []
    
    for po in unique_data_po:
        po_norm = normalize_str(po)
        po_key = po_norm.lower()
        
        # Try direct or clean lookup in JSON mappings
        am_info = None
        if po_key in std_mappings:
            am_info = std_mappings[po_key]
        else:
            # Try fuzzy check in standard keys (e.g. ignoring spacing/casing or matching substring)
            # Standardize and clean key
            clean_po = po_key.replace("bưu cục", "").replace("bc", "").replace("bưu cục", "").replace(" ", "")
            for k, v in std_mappings.items():
                clean_k = k.replace("bưu cục", "").replace("bc", "").replace("bưu cục", "").replace(" ", "")
                if clean_po == clean_k or clean_po in clean_k or clean_k in clean_po:
                    am_info = v
                    break
                    
        if am_info:
            matched += 1
            print(f"Matched: '{po}' -> AM: {am_info[0]}, Province: {am_info[1]}")
        else:
            unmatched.append(po)
            
    print(f"\nMatched {matched} / {len(unique_data_po)}")
    print(f"Unmatched count: {len(unmatched)}")
    for u in unmatched:
        print(f"  - '{u}'")

if __name__ == "__main__":
    main()
