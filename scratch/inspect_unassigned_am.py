import os
import sys
import io
import gspread
import pandas as pd
from google.oauth2.service_account import Credentials
import unicodedata
import json

os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except AttributeError:
    pass

BASE_DIR = r"c:\Users\lap4all\Documents\Auto report"
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

def get_am_province_mappings():
    json_path = os.path.join(BASE_DIR, "scratch", "extracted_mappings.json")
    if os.path.exists(json_path):
        try:
            with open(json_path, "r", encoding="utf-8") as f:
                mappings = json.load(f)
            return {normalize_str(k).lower(): v for k, v in mappings.items()}
        except Exception as e:
            print(f"⚠️ Warning loading extracted_mappings.json: {e}")
    return {}

def resolve_po_info(po_name, cocau_map, std_mappings):
    po_norm = normalize_str(po_name)
    po_key = po_norm.lower()
    if po_key in cocau_map:
        return cocau_map[po_key]
    if po_key in std_mappings:
        return std_mappings[po_key]
    clean_po = po_key.replace("bưu cục", "").replace("bc", "").replace(" ", "")
    for k, v in std_mappings.items():
        clean_k = k.replace("bưu cục", "").replace("bc", "").replace(" ", "")
        if clean_po == clean_k or clean_po in clean_k or clean_k in clean_po:
            return v
    found_cocau = []
    for bc_name, info in cocau_map.items():
        clean_bc = clean_bc_name(bc_name)
        if clean_bc and clean_bc in po_key:
            found_cocau.append((clean_bc, info))
    if found_cocau:
        found_cocau.sort(key=lambda x: len(x[0]), reverse=True)
        return found_cocau[0][1]
    return ("Chưa gán AM", "Chưa gán Tỉnh")

def main():
    credentials = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(credentials)
    sh = gc_client.open_by_key(SHEET_KEY)
    
    # Read Cơ cấu tab
    ws_cocau = sh.worksheet("Cơ cấu")
    cocau_rows = ws_cocau.get_all_values()
    df_cocau = pd.DataFrame(cocau_rows[1:], columns=cocau_rows[0])
    
    cocau_map = {}
    for idx, row in df_cocau.iterrows():
        bc_name = normalize_str(row['Bưu cục'])
        cocau_map[bc_name.lower()] = (row['AM'], row['Tỉnh'])
        
    std_mappings = get_am_province_mappings()
    
    # Read data rớt LC tab
    ws_rot = sh.worksheet("data rớt LC")
    rot_rows = ws_rot.get_all_values()
    if len(rot_rows) < 2:
        print("data rớt LC sheet has no data rows.")
        return
        
    df_rot = pd.DataFrame(rot_rows[1:], columns=rot_rows[0])
    
    print(f"Total rows in 'data rớt LC': {len(df_rot)}")
    
    unassigned_pos = set()
    for idx, row in df_rot.iterrows():
        po_name = row['tenbcxuat']
        if not po_name:
            continue
        am, prov = resolve_po_info(po_name, cocau_map, std_mappings)
        if am == "Chưa gán AM":
            unassigned_pos.add(po_name)
            
    if unassigned_pos:
        print("\n❌ Found bưu cục (tenbcxuat) that are unassigned (Chưa gán AM):")
        for po in sorted(unassigned_pos):
            print(f"- '{po}'")
    else:
        print("\n✅ All bưu cục in 'data rớt LC' are correctly assigned to an AM!")

if __name__ == "__main__":
    main()
