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
    # Standardize string for comparison: NFC normalization, lowercase, strip
    return unicodedata.normalize('NFC', str(s).strip().lower())

def clean_bc_name(name):
    name = normalize_str(name)
    # Remove prefix tags like (dno), (ldo), etc.
    for tag in ['(dno)', '(ldo)', '(kho)', '(bth)', '(nth)']:
        name = name.replace(tag, "")
    # Remove common operational prefixes
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
    
    cocau_list = []
    for idx, row in df_cocau.iterrows():
        orig_bc = row['Bưu cục']
        clean_bc = clean_bc_name(orig_bc)
        cocau_list.append({
            'orig_bc': orig_bc,
            'clean_bc': clean_bc,
            'province': row['Tỉnh'],
            'am': row['AM']
        })
        
    # Custom mapping dictionary for difficult ones as defined in summary
    custom_rules = {
        'đắk mil': ('trần thị nhung', 'Đắk Nông'),
        'gia nghĩa': ('trần văn phước', 'Đắk Nông'),
        'ninh hải': ('nguyễn duy long', 'Ninh Thuận'),
        'lâm hà': ('huỳnh thị kim chi', 'Lâm Đồng'),
        'đạ teh': ('nguyễn lê nguyên vũ', 'Lâm Đồng'),
        'đạ tẻh': ('nguyễn lê nguyên vũ', 'Lâm Đồng'),
        'lạc dương': ('lê minh đại', 'Lâm Đồng'),
        'langbiang': ('lê minh đại', 'Lâm Đồng'),
        'ninh hoà': ('phạm bá thanh công', 'Khánh Hòa'),
        'ninh hòa': ('phạm bá thanh công', 'Khánh Hòa'),
        'cam lâm': ('nguyễn hoàng phi', 'Khánh Hòa'),
    }
    
    matched_count = 0
    unmatched_pos = []
    
    for po in unique_data_po:
        po_norm = normalize_str(po)
        po_clean = clean_bc_name(po)
        
        assigned_am = None
        assigned_prov = None
        
        # 1. Try custom rules first
        for key, (am, prov) in custom_rules.items():
            # If Đà Lạt, apply special rule
            if 'đà lạt' in po_norm:
                if '300 tự phước' in po_norm:
                    assigned_am = 'Lê Văn Trường'
                    assigned_prov = 'Lâm Đồng'
                else:
                    assigned_am = 'Lê Minh Đại'
                    assigned_prov = 'Lâm Đồng'
                break
            elif key in po_norm:
                assigned_am = am
                assigned_prov = prov
                break
                
        if assigned_am:
            matched_count += 1
            print(f"Mapped via CUSTOM RULE: '{po}' -> {assigned_am} ({assigned_prov})")
            continue
            
        # 2. Try matching based on clean_bc being in po_norm
        found_matches = []
        for c in cocau_list:
            if c['clean_bc'] and c['clean_bc'] in po_norm:
                found_matches.append(c)
                
        if len(found_matches) == 1:
            matched_count += 1
            assigned_am = found_matches[0]['am']
            assigned_prov = found_matches[0]['province']
            print(f"Mapped via Clean Substring: '{po}' -> {assigned_am} ({assigned_prov}) [Matched '{found_matches[0]['orig_bc']}']")
        elif len(found_matches) > 1:
            # Ambiguity, select the longest clean substring match to be specific
            found_matches.sort(key=lambda x: len(x['clean_bc']), reverse=True)
            matched_count += 1
            assigned_am = found_matches[0]['am']
            assigned_prov = found_matches[0]['province']
            print(f"Mapped via Ambiguous Clean Substring (selected longest): '{po}' -> {assigned_am} ({assigned_prov}) [Matched '{found_matches[0]['orig_bc']}']")
        else:
            unmatched_pos.append(po)
            
    print(f"\nTotal unique POs in Data: {len(unique_data_po)}")
    print(f"Matched successfully: {matched_count} / {len(unique_data_po)}")
    print(f"Unmatched: {len(unmatched_pos)}")
    for u in unmatched_pos:
        print(f"  - '{u}'")

if __name__ == "__main__":
    main()
