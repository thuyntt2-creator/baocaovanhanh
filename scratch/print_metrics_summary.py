import os
import sys
import io
import gspread
import pandas as pd
from google.oauth2.service_account import Credentials
import unicodedata

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
    import json
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
    
    ws_cocau = sh.worksheet("Cơ cấu")
    cocau_rows = ws_cocau.get_all_values()
    df_cocau = pd.DataFrame(cocau_rows[1:], columns=cocau_rows[0])
    
    cocau_map = {}
    for idx, row in df_cocau.iterrows():
        bc_name = normalize_str(row['Bưu cục'])
        cocau_map[bc_name.lower()] = (row['AM'], row['Tỉnh'])
        
    std_mappings = get_am_province_mappings()
    
    date_str = "2026-06-26"
    categories = ['Khác', 'Shopee', 'TTS']
    
    dfs_updated = {}
    for cat in categories:
        ws_data = sh.worksheet(f"Data {cat}")
        data_values = ws_data.get_all_values()
        df_data = pd.DataFrame(data_values[1:], columns=data_values[0])
        df_data['Vol cần LC'] = pd.to_numeric(df_data['Vol cần LC'], errors='coerce').fillna(0)
        df_data['%_rot_lc'] = pd.to_numeric(df_data['%_rot_lc'], errors='coerce').fillna(0)
        df_data['Vol rớt LC'] = df_data['Vol cần LC'] * df_data['%_rot_lc']
        dfs_updated[cat] = df_data

    # Calculate metrics (excluding 'Chưa gán AM')
    metrics_summary = {}
    n2_date_str = "2026-06-25"
    has_n2 = True
    
    for cat in categories:
        df_n1_cat = dfs_updated[cat][dfs_updated[cat]['Loại ngày'] == date_str].copy()
        df_n1_cat['AM_info'] = df_n1_cat['Chi tiết'].apply(lambda x: resolve_po_info(x, cocau_map, std_mappings))
        df_n1_cat['AM_name'] = df_n1_cat['AM_info'].apply(lambda x: x[0])
        df_n1_cat = df_n1_cat[df_n1_cat['AM_name'] != 'Chưa gán AM']
        
        n1_can = df_n1_cat['Vol cần LC'].sum()
        n1_rot = df_n1_cat['Vol rớt LC'].sum()
        n1_rate = (n1_rot / n1_can * 100) if n1_can > 0 else 0.0
        
        df_n2_cat = dfs_updated[cat][dfs_updated[cat]['Loại ngày'] == n2_date_str].copy()
        df_n2_cat['AM_info'] = df_n2_cat['Chi tiết'].apply(lambda x: resolve_po_info(x, cocau_map, std_mappings))
        df_n2_cat['AM_name'] = df_n2_cat['AM_info'].apply(lambda x: x[0])
        df_n2_cat = df_n2_cat[df_n2_cat['AM_name'] != 'Chưa gán AM']
        
        n2_can = df_n2_cat['Vol cần LC'].sum()
        n2_rot = df_n2_cat['Vol rớt LC'].sum()
        n2_rate = (n2_rot / n2_can * 100) if n2_can > 0 else 0.0
        
        diff = n1_rate - n2_rate
        diff_str = "không đổi"
        if diff > 0:
            diff_str = f"tăng +{diff:.2f}%"
        elif diff < 0:
            diff_str = f"giảm {diff:.2f}%"
        diff_msg = f"(<b>{diff_str}</b> so với hôm qua)"
        
        metrics_summary[cat] = {
            'n1_can': n1_can,
            'n1_rot': n1_rot,
            'n1_rate': n1_rate,
            'n2_can': n2_can,
            'n2_rot': n2_rot,
            'n2_rate': n2_rate,
            'diff_msg': diff_msg
        }
        
    n1_total_can = sum(metrics_summary[cat]['n1_can'] for cat in categories)
    n1_total_rot = sum(metrics_summary[cat]['n1_rot'] for cat in categories)
    n1_total_rate = (n1_total_rot / n1_total_can * 100) if n1_total_can > 0 else 0.0
    
    n2_total_can = sum(metrics_summary[cat]['n2_can'] for cat in categories)
    n2_total_rot = sum(metrics_summary[cat]['n2_rot'] for cat in categories)
    n2_total_rate = (n2_total_rot / n2_total_can * 100) if n2_total_can > 0 else 0.0
    
    diff_tot = n1_total_rate - n2_total_rate
    diff_tot_str = "không đổi"
    if diff_tot > 0:
        diff_tot_str = f"tăng +{diff_tot:.2f}%"
    elif diff_tot < 0:
        diff_tot_str = f"giảm {diff_tot:.2f}%"
    diff_tot_msg = f"(<b>{diff_tot_str}</b> so với hôm qua)"
    
    metrics_summary['Total'] = {
        'n1_can': n1_total_can,
        'n1_rot': n1_total_rot,
        'n1_rate': n1_total_rate,
        'diff_msg': diff_tot_msg
    }
    
    print("\nCALCULATED METRICS (EXCLUDING CHƯA GÁN AM):")
    print(f"Total: can={n1_total_can}, rot={n1_total_rot}, rate={n1_total_rate:.2f}% {diff_tot_msg}")
    for cat in categories:
        m = metrics_summary[cat]
        print(f"{cat}: can={m['n1_can']}, rot={m['n1_rot']}, rate={m['n1_rate']:.2f}% {m['diff_msg']}")

if __name__ == "__main__":
    main()
