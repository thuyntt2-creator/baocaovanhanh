import os
import sys
import io
import json
import pandas as pd
import numpy as np
import gspread
import unicodedata
from google.oauth2.service_account import Credentials
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
SHEET_KEY = '14r8n9L2cIG1Bmz8kSH79B24QzmnOApZhniGyOU40hr4'
scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

def normalize_str(s):
    if not s: return ""
    return unicodedata.normalize('NFC', str(s).strip())

def clean_bc_name(name):
    name = normalize_str(name).lower()
    for tag in ['(dno)', '(ldo)', '(kho)', '(bth)', '(nth)']:
        name = name.replace(tag, "")
    for prefix in ['kho chuyển tiếp', 'kho trung chuyển', 'điểm xử lý hàng', 'điểm lấy hàng', 'bưu cục', 'bc', 'đl']:
        name = name.replace(prefix, "")
    return name.strip()

def get_am_province_mappings():
    json_path = os.path.join(BASE_DIR, "extracted_mappings.json")
    if os.path.exists(json_path):
        with open(json_path, "r", encoding="utf-8") as f:
            return {normalize_str(k).lower(): v for k, v in json.load(f).items()}
    return {}

def resolve_po_info(po_name, cocau_map, std_mappings):
    po_key = normalize_str(po_name).lower()
    if po_key in cocau_map: return cocau_map[po_key]
    if po_key in std_mappings: return std_mappings[po_key]
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

gc_client = gspread.authorize(Credentials.from_service_account_file(JSON_FILE, scopes=scopes))
sh = gc_client.open_by_key(SHEET_KEY)

ws_cocau = None
for sname in ["CoCauVung", "Cơ cấu", "cơ cấu"]:
    try:
        ws_cocau = sh.worksheet(sname)
        break
    except Exception:
        pass
if not ws_cocau:
    print("Worksheet not found")
    sys.exit(1)
cocau_rows = ws_cocau.get_all_values()
df_cocau = pd.DataFrame(cocau_rows[1:], columns=cocau_rows[0])
cocau_map = {normalize_str(row['Bưu cục']).lower(): (row['AM'], row['Tỉnh']) for _, row in df_cocau.iterrows()}
std_mappings = get_am_province_mappings()

ws_tts = sh.worksheet("TTS")
tts_vals = ws_tts.get_all_values()
df_tts = pd.DataFrame(tts_vals[1:], columns=tts_vals[0])

latest_date = df_tts['ltc_date'].max()
tts_latest = df_tts[df_tts['ltc_date'] == latest_date].copy()
tts_latest['vol_ltc'] = pd.to_numeric(tts_latest['vol_ltc'], errors='coerce').fillna(0)
tts_latest['ontime_xuat_first_mile'] = pd.to_numeric(tts_latest['ontime_xuat_first_mile'], errors='coerce').fillna(0)

# Aggregation by bc_lay
tts_latest = tts_latest.groupby('bc_lay', as_index=False).agg({
    'vol_ltc': 'sum',
    'ontime_xuat_first_mile': 'sum'
})
tts_latest['rot_count'] = tts_latest['vol_ltc'] - tts_latest['ontime_xuat_first_mile']
tts_latest['rate'] = np.where(tts_latest['vol_ltc'] > 0, tts_latest['rot_count'] / tts_latest['vol_ltc'], 0.0)

top10 = tts_latest.sort_values(by='rot_count', ascending=False).head(10).reset_index(drop=True)

print(f"=== TOP 10 BƯU CỤC RỚT LC TTS (GỘP BƯU CỤC) NGÀY {latest_date} ===")
for idx, r in top10.iterrows():
    po_name = r['bc_lay']
    am, _ = resolve_po_info(po_name, cocau_map, std_mappings)
    print(f"{idx+1:2d}. {po_name:<30} | AM: {am:<25} | Cần: {r['vol_ltc']:4.0f} | Rớt: {r['rot_count']:4.0f} | Tỷ lệ: {r['rate']*100:6.2f}%")
