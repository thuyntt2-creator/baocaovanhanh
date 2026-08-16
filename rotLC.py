import sys
import io
import os
import re
import json
import asyncio
import pandas as pd
import numpy as np
import gspread
import unicodedata
import requests
from datetime import datetime, timedelta
from google.oauth2.service_account import Credentials
from playwright.sync_api import sync_playwright
from dotenv import load_dotenv

# Fix encoding for Task Scheduler / Command Prompt
os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace', write_through=True)
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace', write_through=True)
except AttributeError:
    pass


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
SHEET_KEY = '14r8n9L2cIG1Bmz8kSH79B24QzmnOApZhniGyOU40hr4'
GTALK_OA_TOKEN = os.environ.get("ROT_LC_GTALK_OA_TOKEN") or "2067164759497973760:RfgqBJY4QtV18udu1wMEfhmoRNI4hgBv"
GTALK_CHANNEL_ID = os.environ.get("ROT_LC_GTALK_CHANNEL_ID") or "2073026093506625536"

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def parse_percentage(val):
    if pd.isna(val) or val == "":
        return 0.0
    val_str = str(val).strip()
    if val_str.endswith('%'):
        try:
            return float(val_str.rstrip('%')) / 100.0
        except ValueError:
            return 0.0
    try:
        return float(val_str)
    except ValueError:
        return 0.0

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
    """Extract standard PO -> (AM, Province) mapping from extracted_mappings.json"""
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
    
    # 1. Direct match in Cơ cấu
    if po_key in cocau_map:
        return cocau_map[po_key]
        
    # 2. Direct match in JSON mappings
    if po_key in std_mappings:
        return std_mappings[po_key]
        
    # 3. Fuzzy match in JSON mappings (substring check)
    clean_po = po_key.replace("bưu cục", "").replace("bc", "").replace("bưu cục", "").replace(" ", "")
    for k, v in std_mappings.items():
        clean_k = k.replace("bưu cục", "").replace("bc", "").replace("bưu cục", "").replace(" ", "")
        if clean_po == clean_k or clean_po in clean_k or clean_k in clean_po:
            return v
            
    # 4. Fuzzy clean match in Cơ cấu (longest clean match)
    found_cocau = []
    for bc_name, info in cocau_map.items():
        clean_bc = clean_bc_name(bc_name)
        if clean_bc and clean_bc in po_key:
            found_cocau.append((clean_bc, info))
            
    if found_cocau:
        found_cocau.sort(key=lambda x: len(x[0]), reverse=True)
        return found_cocau[0][1]
        
    return ("Chưa gán AM", "Chưa gán Tỉnh")

def get_best_master_po(raw_po, po_list):
    raw_norm = normalize_str(raw_po).lower()
    # 1. Exact match
    for p in po_list:
        if normalize_str(p).lower() == raw_norm:
            return p
    # 2. Clean exact match
    raw_clean = clean_bc_name(raw_po)
    if not raw_clean:
        return None
    for p in po_list:
        if clean_bc_name(p) == raw_clean:
            return p
    # 3. Substring match
    matches = []
    for p in po_list:
        p_clean = clean_bc_name(p)
        if p_clean and (raw_clean in p_clean or p_clean in raw_clean):
            matches.append(p)
    if matches:
        matches.sort(key=lambda x: abs(len(clean_bc_name(x)) - len(raw_clean)))
        return matches[0]
    return None

def main():
    # 1. Setup target dates
    # Target date default: yesterday (N-1)
    target_date = datetime.now() - timedelta(days=1)
    date_str = target_date.strftime("%Y-%m-%d")
    print(f"📅 Target date is N-1: {date_str}")
    
    # 2. Read spreadsheet and map sheets
    credentials = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(credentials)
    sh = gc_client.open_by_key(SHEET_KEY)
    
    ws_cocau = None
    for sname in ["CoCauVung", "Cơ cấu", "cơ cấu"]:
        try:
            ws_cocau = sh.worksheet(sname)
            break
        except Exception:
            pass
    if not ws_cocau:
        print("❌ Không tìm thấy tab CoCauVung hoặc Cơ cấu trong spreadsheet.")
        sys.exit(1)
        
    cocau_rows = ws_cocau.get_all_values()
    df_cocau = pd.DataFrame(cocau_rows[1:], columns=cocau_rows[0])
    
    # Build Cơ cấu maps
    cocau_map = {}
    for idx, row in df_cocau.iterrows():
        bc_name = normalize_str(row['Bưu cục'])
        cocau_map[bc_name.lower()] = (row['AM'], row['Tỉnh'])
        
    std_mappings = get_am_province_mappings()
    
    # 3. Read sheet "TTS" from Google Sheets (Source of truth for TTS volume & ontime)
    print("📖 Reading 'TTS' worksheet from Google Sheets for TTS calculations...")
    ws_tts_sheet = sh.worksheet("TTS")
    tts_vals = ws_tts_sheet.get_all_values()
    df_tts_sheet = pd.DataFrame(tts_vals[1:], columns=tts_vals[0])
    
    # Filter only shift == 'Tối' (exclude 'Ngoài giờ cutoff')
    if 'shift' in df_tts_sheet.columns:
        print("🔍 Filtering TTS sheet data: keeping only shift == 'Tối' (excluding 'Ngoài giờ cutoff')...")
        df_tts_sheet = df_tts_sheet[df_tts_sheet['shift'].astype(str).str.strip().str.lower() == 'tối'].copy()
    
    n2_date_str = (target_date - timedelta(days=1)).strftime("%Y-%m-%d")
    
    # Filter by ltc_date if available
    if 'ltc_date' in df_tts_sheet.columns:
        tts_latest = df_tts_sheet[df_tts_sheet['ltc_date'] == date_str].copy()
        if len(tts_latest) == 0:
            latest_available_date = df_tts_sheet['ltc_date'].max()
            print(f"⚠️ No rows for date {date_str} in TTS sheet, falling back to latest date: {latest_available_date}")
            tts_latest = df_tts_sheet[df_tts_sheet['ltc_date'] == latest_available_date].copy()
        tts_n2 = df_tts_sheet[df_tts_sheet['ltc_date'] == n2_date_str].copy()
    else:
        start_indices = df_tts_sheet[df_tts_sheet['bc_lay'].str.contains('Bắc Bình', case=False, na=False)].index
        if len(start_indices) > 0:
            last_start = start_indices[-1]
            tts_latest = df_tts_sheet.iloc[last_start:].copy()
            tts_n2 = df_tts_sheet.iloc[start_indices[-2]:last_start].copy() if len(start_indices) >= 2 else pd.DataFrame()
        else:
            tts_latest = df_tts_sheet.tail(100).copy()
            tts_n2 = pd.DataFrame()
        
    tts_latest['vol_ltc'] = pd.to_numeric(tts_latest['vol_ltc'], errors='coerce').fillna(0)
    tts_latest['ontime_xuat_first_mile'] = pd.to_numeric(tts_latest['ontime_xuat_first_mile'], errors='coerce').fillna(0)
    
    # Aggregate by bc_lay so that multiple shifts / phân nhóm (cột E) for the same PO are combined
    tts_latest = tts_latest.groupby('bc_lay', as_index=False).agg({
        'vol_ltc': 'sum',
        'ontime_xuat_first_mile': 'sum'
    })
    tts_latest['rot_count'] = tts_latest['vol_ltc'] - tts_latest['ontime_xuat_first_mile']
    tts_latest['rate'] = np.where(tts_latest['vol_ltc'] > 0, tts_latest['rot_count'] / tts_latest['vol_ltc'], 0.0)
    
    if len(tts_n2) > 0:
        tts_n2['vol_ltc'] = pd.to_numeric(tts_n2['vol_ltc'], errors='coerce').fillna(0)
        tts_n2['ontime_xuat_first_mile'] = pd.to_numeric(tts_n2['ontime_xuat_first_mile'], errors='coerce').fillna(0)
        tts_n2 = tts_n2.groupby('bc_lay', as_index=False).agg({
            'vol_ltc': 'sum',
            'ontime_xuat_first_mile': 'sum'
        })
        tts_n2['rot_count'] = tts_n2['vol_ltc'] - tts_n2['ontime_xuat_first_mile']
        tts_n2['rate'] = np.where(tts_n2['vol_ltc'] > 0, tts_n2['rot_count'] / tts_n2['vol_ltc'], 0.0)
    
    # Build master PO mapping for TTS sheet
    tts_po_map = {}
    for idx, row in tts_latest.iterrows():
        tts_po_map[row['bc_lay']] = (row['vol_ltc'], row['rot_count'], row['rate'])

    # 4. Update 'Data TTS' worksheet
    print("📊 Calculating %_rot_lc and updating 'Data TTS' worksheet...")
    ws_data = sh.worksheet("Data TTS")
    data_values = ws_data.get_all_values()
    df_data = pd.DataFrame(data_values[1:], columns=data_values[0])
    
    new_data_rows = [data_values[0]] # Header
    for idx, row in df_data.iterrows():
        row_list = data_values[idx + 1]
        if row['Loại ngày'] == date_str:
            data_po = row['Chi tiết']
            # Match master PO from tts_po_map
            master_po = get_best_master_po(data_po, list(tts_po_map.keys()))
            if master_po and master_po in tts_po_map:
                vol_ltc_val, rot_count_val, rate_val = tts_po_map[master_po]
                row_list[3] = int(vol_ltc_val)
                row_list[4] = float(rate_val)
            else:
                row_list[4] = 0.0
        new_data_rows.append(row_list)
        
    print("📤 Pushing updated Data TTS back to Google Sheets...")
    ws_data.update(new_data_rows, value_input_option='USER_ENTERED')
    print("✅ Data TTS updated.")
    
    # 4.5 Ensure formula in 'data rớt LC' and split by AM
    print("📖 Checking and preserving formula in 'data rớt LC' worksheet...")
    ws_rot = sh.worksheet("data rớt LC")
    rot_a1_formula = ws_rot.acell('A1', value_render_option='FORMULA').value
    expected_formula = '=QUERY(IMPORTRANGE("1LdnXYFnkACLUnVK2-I7MQ8YGRRNc88RcXjDu-lbQQ2Q", "NTB!I1:T"), "SELECT * WHERE Col1 IS NOT NULL", 1)'
    
    if not str(rot_a1_formula).startswith('=QUERY'):
        print("✍️ Restoring formula in A1 of 'data rớt LC'...")
        ws_rot.clear()
        ws_rot.update_acell('A1', expected_formula)
        print("✅ Formula restored.")
    else:
        print("✅ Formula in A1 of 'data rớt LC' is present.")
        
    rot_vals = ws_rot.get_all_values()
    df_raw_ntb_late = pd.DataFrame(rot_vals[1:], columns=rot_vals[0]) if len(rot_vals) > 1 else pd.DataFrame(columns=rot_vals[0] if rot_vals else [])
    print(f"✔️ Sheet 'data rớt LC' currently has {len(df_raw_ntb_late)} rows.")
    
    print("✂️ Splitting data rớt LC by AM...")
    all_ams = df_cocau['AM'].unique()
    all_ams = [str(x).strip() for x in all_ams if x and str(x).strip()]
    
    if len(df_raw_ntb_late) > 0:
        df_raw_ntb_copy = df_raw_ntb_late.copy()
        raw_ams = []
        for idx_e, row_e in df_raw_ntb_copy.iterrows():
            raw_po = row_e.get('tenbcxuat', '') or row_e.get('bc_lay', '')
            am, _ = resolve_po_info(raw_po, cocau_map, std_mappings)
            raw_ams.append(am)
        df_raw_ntb_copy['AM_resolved'] = raw_ams
        
        for am_name in all_ams:
            df_am = df_raw_ntb_copy[df_raw_ntb_copy['AM_resolved'] == am_name].copy()
            df_am = df_am.drop(columns=['AM_resolved'])
            try:
                ws_am = sh.worksheet(am_name)
            except gspread.exceptions.WorksheetNotFound:
                ws_am = sh.add_worksheet(title=am_name, rows=1000, cols=len(df_raw_ntb_late.columns))
                
            ws_am.clear()
            if len(df_am) > 0:
                am_data_to_write = [df_am.columns.values.tolist()] + df_am.values.tolist()
                ws_am.update(am_data_to_write, value_input_option='USER_ENTERED')
            else:
                ws_am.update([df_am.columns.values.tolist()], value_input_option='USER_ENTERED')
    print("✅ AM worksheets updated successfully.")

    # 5. Extract TTS metrics from tts_latest for GTalk message
    n1_tts_can = tts_latest['vol_ltc'].sum()
    n1_tts_rot = tts_latest['rot_count'].sum()
    n1_tts_rate = (n1_tts_rot / n1_tts_can * 100) if n1_tts_can > 0 else 0.0
    
    # Calculate N-2 block comparison
    diff_msg = "(so với hôm qua: N/A)"
    if len(tts_n2) > 0:
        tts_n2['vol_ltc'] = pd.to_numeric(tts_n2['vol_ltc'], errors='coerce').fillna(0)
        tts_n2['ontime_xuat_first_mile'] = pd.to_numeric(tts_n2['ontime_xuat_first_mile'], errors='coerce').fillna(0)
        n2_can = tts_n2['vol_ltc'].sum()
        n2_rot = (tts_n2['vol_ltc'] - tts_n2['ontime_xuat_first_mile']).sum()
        n2_rate = (n2_rot / n2_can * 100) if n2_can > 0 else 0.0
        diff = n1_tts_rate - n2_rate
        diff_str = "không đổi"
        if diff > 0:
            diff_str = f"tăng +{diff:.2f}%"
        elif diff < 0:
            diff_str = f"giảm {diff:.2f}%"
        diff_msg = f"(<b>{diff_str}</b> so với hôm qua)"
        
    # 6. Render visual tables and capture screenshots (TTS Only)
    # Table 1: Top 10 POs by TTS dropped orders count from sheet TTS
    top10_tts = tts_latest.sort_values(by='rot_count', ascending=False).head(10).reset_index(drop=True)
    
    t1_rows_html = ""
    t1_sum_tts = 0
    t1_sum_can = 0
    for idx, row in top10_tts.iterrows():
        po_name = row['bc_lay']
        am, _ = resolve_po_info(po_name, cocau_map, std_mappings)
        vol_can = float(row['vol_ltc'])
        vol_rot = float(row['rot_count'])
        rate = (vol_rot / vol_can * 100) if vol_can > 0 else 0.0
        
        t1_rows_html += f"""
        <tr>
            <td class="number">{idx + 1}</td>
            <td class="po-name">{po_name}</td>
            <td class="am-name">{am}</td>
            <td class="number" style="text-align:center;">{vol_can:,.0f}</td>
            <td class="number grand-total" style="text-align:center;">{vol_rot:,.0f}</td>
            <td class="number" style="text-align:center;">{rate:.2f}%</td>
        </tr>
        """
        t1_sum_tts += vol_rot
        t1_sum_can += vol_can
        
    t1_total_rate = (t1_sum_tts / t1_sum_can) * 100 if t1_sum_can > 0 else 0.0
    t1_rows_html += f"""
    <tr class="total-row">
        <td colspan="3">TỔNG CỘNG TOP 10 (TTS)</td>
        <td class="number" style="text-align:center;">{t1_sum_can:,.0f}</td>
        <td class="number" style="text-align:center;">{t1_sum_tts:,.0f}</td>
        <td class="number" style="text-align:center;">{t1_total_rate:.2f}%</td>
    </tr>
    """
    
    t1_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="utf-8">
    <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <style>
        body {{
            font-family: 'Inter', sans-serif;
            background-color: #f8fafc;
            margin: 0;
            padding: 24px;
            display: flex;
            justify-content: center;
            align-items: center;
        }}
        #capture-container {{
            background: #ffffff;
            padding: 32px;
            border-radius: 16px;
            box-shadow: 0 4px 20px rgba(15, 23, 42, 0.05);
            border: 1px solid #e2e8f0;
            width: 850px;
        }}
        .header {{
            margin-bottom: 24px;
            text-align: center;
            border-bottom: 2px solid #f1f5f9;
            padding-bottom: 16px;
        }}
        .header h2 {{
            margin: 0;
            font-family: 'Plus Jakarta Sans', sans-serif;
            font-size: 24px;
            color: #1e3a8a;
            font-weight: 800;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }}
        .header p {{
            margin: 4px 0 0 0;
            font-size: 14px;
            color: #64748b;
            font-weight: 500;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            text-align: left;
        }}
        th {{
            font-family: 'Plus Jakarta Sans', sans-serif;
            background-color: #1e3a8a;
            color: #ffffff;
            font-weight: 700;
            font-size: 13px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            padding: 10px 8px;
            border-bottom: 2px solid #e2e8f0;
        }}
        td {{
            padding: 9px 8px;
            font-size: 13px;
            color: #334155;
            border-bottom: 1px solid #f1f5f9;
            font-weight: 500;
        }}
        tr:nth-child(even) td {{
            background-color: #f8fafc;
        }}
        .po-name {{
            font-weight: 600;
            color: #0f172a;
        }}
        .am-name {{
            color: #475569;
            font-weight: 500;
        }}
        .number {{
            text-align: center;
            font-weight: 600;
        }}
        .grand-total {{
            font-weight: 700;
            color: #1e3a8a;
        }}
        .total-row td {{
            background-color: #eff6ff !important;
            font-weight: 800;
            color: #1e3a8a;
            border-top: 2px solid #3b82f6;
            border-bottom: none;
            padding: 12px 8px;
            font-size: 13px;
        }}
    </style>
    </head>
    <body>
    <div id="capture-container">
        <div class="header">
            <h2>Báo Cáo Top 10 Bưu Cục Rớt LC TTS</h2>
            <p>Khu vực Nam Trung Bộ | Ngày {target_date.strftime('%d/%m/%Y')}</p>
        </div>
        <table>
            <thead>
                <tr>
                    <th style="width: 6%; text-align:center;">STT</th>
                    <th style="width: 35%;">Chi tiết</th>
                    <th style="width: 23%;">AM</th>
                    <th style="width: 12%; text-align:center;">Vol cần LC</th>
                    <th style="width: 12%; text-align:center;">Vol Rớt</th>
                    <th style="width: 12%; text-align:center;">Tỷ Lệ Rớt</th>
                </tr>
            </thead>
            <tbody>
                {t1_rows_html}
            </tbody>
        </table>
    </div>
    </body>
    </html>
    """
    
    # Table 2: Thống kê rớt LC TTS theo AM
    tts_latest_copy = tts_latest.copy()
    tts_latest_copy['AM_info'] = tts_latest_copy['bc_lay'].apply(lambda x: resolve_po_info(x, cocau_map, std_mappings))
    tts_latest_copy['AM_name'] = tts_latest_copy['AM_info'].apply(lambda x: x[0])
    tts_latest_copy = tts_latest_copy[tts_latest_copy['AM_name'] != 'Chưa gán AM']
    
    am_tts_n1_rates = tts_latest_copy.groupby('AM_name').agg(
        n1_can=('vol_ltc', 'sum'),
        n1_rot=('rot_count', 'sum')
    ).reset_index()
    am_tts_n1_rates['n1_rate'] = am_tts_n1_rates['n1_rot'] / am_tts_n1_rates['n1_can']
    am_tts_n1_rates = am_tts_n1_rates.sort_values(by='n1_rate', ascending=False).reset_index(drop=True)
    
    am_tts_n2_map = {}
    if len(tts_n2) > 0:
        tts_n2_copy = tts_n2.copy()
        tts_n2_copy['vol_ltc'] = pd.to_numeric(tts_n2_copy['vol_ltc'], errors='coerce').fillna(0)
        tts_n2_copy['ontime_xuat_first_mile'] = pd.to_numeric(tts_n2_copy['ontime_xuat_first_mile'], errors='coerce').fillna(0)
        tts_n2_copy['rot_count'] = tts_n2_copy['vol_ltc'] - tts_n2_copy['ontime_xuat_first_mile']
        tts_n2_copy['AM_info'] = tts_n2_copy['bc_lay'].apply(lambda x: resolve_po_info(x, cocau_map, std_mappings))
        tts_n2_copy['AM_name'] = tts_n2_copy['AM_info'].apply(lambda x: x[0])
        am_n2_agg = tts_n2_copy.groupby('AM_name').agg(can=('vol_ltc', 'sum'), rot=('rot_count', 'sum')).reset_index()
        am_n2_agg['rate'] = am_n2_agg['rot'] / am_n2_agg['can']
        am_tts_n2_map = dict(zip(am_n2_agg['AM_name'], am_n2_agg['rate']))
    
    t2_rows_html = ""
    for idx, row in am_tts_n1_rates.iterrows():
        am_name = row['AM_name']
        vol_can = row['n1_can']
        vol_rot = row['n1_rot']
        rate = row['n1_rate'] * 100
        
        if rate < 2.0:
            badge_class = "rate-green"
            dot_class = "dot-green"
        elif rate <= 5.0:
            badge_class = "rate-yellow"
            dot_class = "dot-yellow"
        else:
            badge_class = "rate-red"
            dot_class = "dot-red"
            
        if am_tts_n2_map:
            prev_rate = am_tts_n2_map.get(am_name, 0.0) * 100
            diff = rate - prev_rate
            if prev_rate == 0.0 and vol_rot == 0:
                diff_html = '<span class="compare-same">—</span>'
            elif diff > 0:
                diff_html = f'<span class="compare-up">▲ +{diff:.2f}%</span>'
            elif diff < 0:
                diff_html = f'<span class="compare-down">▼ {diff:.2f}%</span>'
            else:
                diff_html = '<span class="compare-same">—</span>'
        else:
            diff_html = '<span class="compare-same">N/A</span>'
            
        t2_rows_html += f"""
        <tr>
            <td class="number">{idx + 1}</td>
            <td class="am-name">{am_name}</td>
            <td class="number">{vol_can:,.0f}</td>
            <td class="number">{vol_rot:,.0f}</td>
            <td style="text-align:center;">
                <span class="rate-badge {badge_class}"><span class="dot {dot_class}"></span>{rate:.2f}%</span>
            </td>
            <td style="text-align:center;">{diff_html}</td>
        </tr>
        """
        
    t2_grand_badge = "rate-yellow" if n1_tts_rate >= 2.0 and n1_tts_rate <= 5.0 else "rate-red" if n1_tts_rate > 5.0 else "rate-green"
    t2_grand_dot = "dot-yellow" if n1_tts_rate >= 2.0 and n1_tts_rate <= 5.0 else "dot-red" if n1_tts_rate > 5.0 else "dot-green"
    
    diff_msg_tot = diff_msg if am_tts_n2_map else "(so với hôm qua: N/A)"
        
    t2_rows_html += f"""
    <tr class="total-row">
        <td colspan="2">TỔNG CỘNG TTS</td>
        <td class="number">{n1_tts_can:,.0f}</td>
        <td class="number">{n1_tts_rot:,.0f}</td>
        <td style="text-align:center;">
            <span class="rate-badge {t2_grand_badge}"><span class="dot {t2_grand_dot}"></span>{n1_tts_rate:.2f}%</span>
        </td>
        <td style="text-align:center;">—</td>
    </tr>
    """
    
    t2_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="utf-8">
    <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <style>
        body {{
            font-family: 'Inter', sans-serif;
            background-color: #f8fafc;
            margin: 0;
            padding: 24px;
            display: flex;
            justify-content: center;
            align-items: center;
        }}
        #capture-container {{
            background: #ffffff;
            padding: 32px;
            border-radius: 16px;
            box-shadow: 0 4px 20px rgba(15, 23, 42, 0.05);
            border: 1px solid #e2e8f0;
            width: 850px;
        }}
        .header {{
            margin-bottom: 24px;
            text-align: center;
            border-bottom: 2px solid #f1f5f9;
            padding-bottom: 16px;
        }}
        .header h2 {{
            margin: 0;
            font-family: 'Plus Jakarta Sans', sans-serif;
            font-size: 24px;
            color: #b91c1c;
            font-weight: 800;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }}
        .header p {{
            margin: 4px 0 0 0;
            font-size: 14px;
            color: #64748b;
            font-weight: 500;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            text-align: left;
        }}
        th {{
            font-family: 'Plus Jakarta Sans', sans-serif;
            background-color: #b91c1c;
            color: #ffffff;
            font-weight: 700;
            font-size: 13px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            padding: 10px 8px;
            border-bottom: 2px solid #e2e8f0;
        }}
        td {{
            padding: 9px 8px;
            font-size: 13px;
            color: #334155;
            border-bottom: 1px solid #f1f5f9;
            font-weight: 500;
        }}
        tr:nth-child(even) td {{
            background-color: #f8fafc;
        }}
        .am-name {{
            font-weight: 600;
            color: #0f172a;
        }}
        .number {{
            text-align: center;
            font-weight: 600;
        }}
        .rate-badge {{
            display: inline-flex;
            align-items: center;
            gap: 6px;
            padding: 4px 10px;
            border-radius: 9999px;
            font-size: 12px;
            font-weight: 700;
            text-align: center;
        }}
        .rate-green {{
            background-color: #dcfce7;
            color: #15803d;
        }}
        .rate-yellow {{
            background-color: #fef9c3;
            color: #a16207;
        }}
        .rate-red {{
            background-color: #ffe4e6;
            color: #b91c1c;
        }}
        .dot {{
            width: 6px;
            height: 6px;
            border-radius: 50%;
            display: inline-block;
        }}
        .dot-green {{ background-color: #15803d; }}
        .dot-yellow {{ background-color: #a16207; }}
        .dot-red {{ background-color: #b91c1c; }}
        .compare-up {{
            color: #ef4444;
            font-size: 11px;
            font-weight: 700;
        }}
        .compare-down {{
            color: #22c55e;
            font-size: 11px;
            font-weight: 700;
        }}
        .compare-same {{
            color: #94a3b8;
            font-size: 11px;
            font-weight: 700;
        }}
        .total-row td {{
            background-color: #fef2f2 !important;
            font-weight: 800;
            color: #b91c1c;
            border-top: 2px solid #ef4444;
            border-bottom: none;
            padding: 12px 8px;
            font-size: 13px;
        }}
    </style>
    </head>
    <body>
    <div id="capture-container">
        <div class="header">
            <h2>Báo Cáo Tỷ Lệ AM Rớt LC TTS</h2>
            <p>Khu vực Nam Trung Bộ | Ngày {target_date.strftime('%d/%m/%Y')}</p>
        </div>
        <table>
            <thead>
                <tr>
                    <th style="width: 6%; text-align:center;">STT</th>
                    <th style="width: 34%;">AM</th>
                    <th style="width: 17%; text-align:center;">Vol cần LC</th>
                    <th style="width: 17%; text-align:center;">Vol rớt LC</th>
                    <th style="width: 13%; text-align:center;">Tỷ lệ</th>
                    <th style="width: 13%; text-align:center;">So ngày cũ</th>
                </tr>
            </thead>
            <tbody>
                {t2_rows_html}
            </tbody>
        </table>
    </div>
    </body>
    </html>
    """
    
    # Save and screenshot HTMLs using Playwright
    t1_html_path = os.path.join(BASE_DIR, "t1_top10_po.html")
    t2_html_path = os.path.join(BASE_DIR, "t2_am_rates.html")
    
    with open(t1_html_path, "w", encoding="utf-8") as f:
        f.write(t1_html)
    with open(t2_html_path, "w", encoding="utf-8") as f:
        f.write(t2_html)
        
    t1_image_path = os.path.join(BASE_DIR, "top10_po_rot_lc.png")
    t2_image_path = os.path.join(BASE_DIR, "am_rates_rot_lc.png")
    
    print("📸 Rendering and capturing images with Playwright...")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(viewport={"width": 1200, "height": 1000})
        
        # Screenshot Table 1
        page.goto(f"file:///{t1_html_path.replace('\\', '/')}")
        page.wait_for_timeout(1000)
        page.locator("#capture-container").screenshot(path=t1_image_path)
        print(f"✔️ Captured image 1: {t1_image_path}")
        
        # Screenshot Table 2
        page.goto(f"file:///{t2_html_path.replace('\\', '/')}")
        page.wait_for_timeout(1000)
        page.locator("#capture-container").screenshot(path=t2_image_path)
        print(f"✔️ Captured image 2: {t2_image_path}")
        
        browser.close()
        
    # Clean up temp HTML files
    try:
        os.remove(t1_html_path)
        os.remove(t2_html_path)
    except:
        pass
        
    # 7. Post report and images to GTalk (TTS Only)
    print("📡 Posting report and images to GTalk channel...")
    caption = f"""<b>BÁO CÁO RỚT LC TTS NGÀY {target_date.strftime('%d/%m/%Y')}</b>

<b>Chỉ số TTS</b>
- Tổng đơn trễ (chưa LC): <b>{n1_tts_rot:,.0f} đơn</b>
- Tổng đơn cần LC: <b>{n1_tts_can:,.0f} đơn</b>
—&gt; Tỷ lệ rớt LC TTS: <b>{n1_tts_rate:.2f}%</b> <b>{diff_msg}</b>

🔗 Link danh sách chi tiết MVĐ rớt LC: <a href="https://docs.google.com/spreadsheets/d/{SHEET_KEY}/edit#gid=1702024"><b>xem chi tiết</b></a>

<i>Logic dữ liệu:
- Đơn được lấy thành công phải được xuất khỏi bưu cục trước 09:30 cùng ngày</i>"""

    def upload_image_to_gtalk(image_path):
        file_name = os.path.basename(image_path)
        file_size = os.path.getsize(image_path)
        with open(image_path, 'rb') as f:
            file_bytes = f.read()
            
        init_payload = {
            "ChannelId": GTALK_CHANNEL_ID,
            "FileName": file_name,
            "FileSize": str(file_size),
            "MimeType": "image/png",
            "Metadata": json.dumps({"width": 1200, "height": 800}),
            "oaToken": GTALK_OA_TOKEN
        }
        
        resp_init = requests.post("https://mbff.ghn.vn/api/gtalk/initiate-upload", json=init_payload)
        if resp_init.status_code == 200:
            init_data = resp_init.json()
            if init_data.get("errorCode") == "success":
                presigned_url = init_data["data"]["PresignedURL"]
                upload_id = init_data["data"]["UploadId"]
                
                resp_put = requests.put(presigned_url, data=file_bytes, headers={"Content-Type": "image/png"})
                if resp_put.status_code == 200:
                    resp_comp = requests.post("https://mbff.ghn.vn/api/gtalk/complete-upload", json={"oaToken": GTALK_OA_TOKEN, "UploadId": upload_id})
                    if resp_comp.status_code == 200:
                        comp_data = resp_comp.json()
                        if comp_data.get("errorCode") == "success":
                            return comp_data["data"]["Id"]
                        else:
                            print(f"⚠️ GTalk complete-upload logic error: {comp_data}")
                    else:
                        print(f"⚠️ GTalk complete-upload HTTP error {resp_comp.status_code}: {resp_comp.text}")
                else:
                    print(f"⚠️ GTalk put-file HTTP error {resp_put.status_code}: {resp_put.text}")
            else:
                print(f"⚠️ GTalk initiate-upload logic error: {init_data}")
        else:
            print(f"⚠️ GTalk initiate-upload HTTP error {resp_init.status_code}: {resp_init.text}")
        return None

    # Upload both images
    img1_id = upload_image_to_gtalk(t1_image_path)
    img2_id = upload_image_to_gtalk(t2_image_path)
    
    if img1_id and img2_id:
        send_payload = {
            "channelId": GTALK_CHANNEL_ID,
            "clientMsgId": str(int(datetime.now().timestamp() * 1000)),
            "content": {
                "parseMode": "HTML",
                "attachment": {
                    "caption": caption,
                    "items": [
                        {"image": {"fileId": img1_id, "width": 1200, "height": 800}},
                        {"image": {"fileId": img2_id, "width": 1200, "height": 800}}
                    ]
                }
            },
            "oaToken": GTALK_OA_TOKEN
        }
        r_send = requests.post("https://mbff.ghn.vn/api/gtalk/send-message", json=send_payload)
        if r_send.status_code == 200 and r_send.json().get("errorCode") == "success":
            print("✅ Successfully posted report message and images to GTalk!")
        else:
            print(f"❌ Failed to send GTalk message: {r_send.text}")
    else:
        print("❌ Image upload to GTalk failed.")

    # Clean up screenshots
    try:
        os.remove(t1_image_path)
        os.remove(t2_image_path)
    except:
        pass
        
    print("🎉 Report calculation and GTalk post completed successfully!")
        
    print("🎉 Report calculation and GTalk post completed successfully!")

if __name__ == "__main__":
    main()
