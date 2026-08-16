import os
import sys
import io
import re
import json
import pandas as pd
import gspread
import unicodedata
import requests
from datetime import datetime
from google.oauth2.service_account import Credentials
from playwright.sync_api import sync_playwright
from dotenv import load_dotenv

# Fix encoding for Windows console logging
sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
SHEET_KEY = '1DAwY-46twFrHIs77R4p4IMuIZ6JTE-e58Aj-9Kcr5Jk'

# Load environment configuration if available
env_path = r"c:\Users\lap4all\Desktop\New folder\.env"
if os.path.exists(env_path):
    load_dotenv(dotenv_path=env_path, override=True)
else:
    load_dotenv()

# Override GTalk credentials from env if defined, otherwise fallback
GTALK_OA_TOKEN = os.environ.get("GTC_CA1_GTALK_OA_TOKEN") or "2067164759497973760:RfgqBJY4QtV18udu1wMEfhmoRNI4hgBv"
GTALK_CHANNEL_ID = os.environ.get("GTC_CA1_GTALK_CHANNEL_ID") or os.environ.get("GTALK_CHANNEL_ID") or "2073028340427608064"

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

def resolve_po_info(po_name, cocau_map):
    po_key = normalize_str(po_name).lower()
    
    # 1. Direct match
    if po_key in cocau_map:
        return cocau_map[po_key]
        
    # 2. Cleaned direct match
    clean_key = clean_bc_name(po_name)
    for k, v in cocau_map.items():
        if clean_bc_name(k) == clean_key:
            return v
            
    # 3. Fuzzy match (substring)
    for k, v in cocau_map.items():
        clean_k = clean_bc_name(k)
        if clean_key and clean_k and (clean_key in clean_k or clean_k in clean_key):
            return v
            
    # Fallback to defaults
    if "dno" in po_key or "đắk nông" in po_key:
        return ("Trần Văn Phước", "Đắk Nông")
    if "ldo" in po_key or "lâm đồng" in po_key:
        return ("Nguyễn Lê Nguyên Vũ", "Lâm Đồng")
    if "bth" in po_key or "bình thuận" in po_key:
        return ("Nguyễn Ngọc Khánh", "Bình Thuận")
    return ("Chưa gán AM", "Chưa gán Tỉnh")

def clean_pct(val):
    if not val or pd.isna(val):
        return 0.0
    val_str = str(val).strip().replace('%', '')
    val_str = val_str.replace(',', '.')
    try:
        num = float(val_str)
        if num > 1.0:
            return num / 100.0
        return num
    except ValueError:
        return 0.0

def clean_volume(val):
    if not val or pd.isna(val):
        return 0
    val_str = str(val).strip().replace('.', '').replace(',', '')
    try:
        return int(round(float(val_str)))
    except ValueError:
        return 0

def main():
    print("🚀 BẮT ĐẦU CHẠY BÁO CÁO GTC CA 1 TTS LÚC:", datetime.now().strftime('%H:%M:%S'))
    
    # 1. Authorize Google Sheet
    print("📡 Đang kết nối tới Google Sheets...")
    oauth_file = os.path.join(BASE_DIR, 'credentials_oauth.json')
    if os.path.exists(oauth_file):
        print("🔑 Sử dụng xác thực OAuth 2.0 (credentials_oauth.json)...")
        gc_client = gspread.oauth(
            credentials_filename=oauth_file,
            authorized_user_filename=os.path.join(BASE_DIR, 'authorized_user.json')
        )
    else:
        credentials = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
        gc_client = gspread.authorize(credentials)
        
    sh = gc_client.open_by_key(SHEET_KEY)
    
    # 2. Read Worksheets
    print("📥 Đang tải các worksheet dữ liệu...")
    ws_cc = None
    for sname in ["CoCauVung", "Cơ cấu", "cơ cấu"]:
        try:
            ws_cc = sh.worksheet(sname)
            break
        except Exception:
            pass
    if not ws_cc:
        print("❌ Không tìm thấy tab CoCauVung hoặc Cơ cấu trong spreadsheet.")
        sys.exit(1)
        
    df_cocau_raw = pd.DataFrame(ws_cc.get_all_records())
    df_gtc_raw = pd.DataFrame(sh.worksheet("rawGTCTTS").get_all_records())
    
    if df_gtc_raw.empty:
        print("❌ Dữ liệu rawGTCTTS trống. Bỏ qua chạy báo cáo.")
        sys.exit(1)
        
    # 3. Build Cơ cấu map
    cocau_map = {}
    
    # Detect AM and Province columns robustly
    am_col = next((c for c in df_cocau_raw.columns if c.lower() == 'am'), 'Am')
    tinh_col = next((c for c in df_cocau_raw.columns if c.lower() == 'tỉnh'), 'Tỉnh')
    bc_col = next((c for c in df_cocau_raw.columns if c.lower() == 'bưu cục'), 'Bưu cục')
    
    for idx, row in df_cocau_raw.iterrows():
        bc_name = normalize_str(row[bc_col]).lower()
        cocau_map[bc_name] = (row[am_col], row[tinh_col])
        
    # 4. Parse Dates and identify latest date (N-1) and comparison date (N-2)
    df_gtc_raw['date_parsed'] = df_gtc_raw['Time'].apply(lambda x: str(x).split(" - ")[0].strip())
    unique_dates = sorted(df_gtc_raw['date_parsed'].unique())
    
    if len(unique_dates) == 0:
        print("❌ Không tìm thấy ngày nào trong rawGTCTTS.")
        sys.exit(1)
        
    latest_date_str = unique_dates[-1]
    prev_date_str = unique_dates[-2] if len(unique_dates) > 1 else None
    
    latest_date_obj = datetime.strptime(latest_date_str, "%Y-%m-%d")
    latest_date_display = latest_date_obj.strftime("%d/%m/%Y")
    
    print(f"📅 Target Date (N-1): {latest_date_str} (Hiển thị: {latest_date_display})")
    if prev_date_str:
        print(f"📅 Comparison Date (N-2): {prev_date_str}")
        
    # Clean numeric data in rawGTCTTS
    df_gtc_raw['Vol_Clean'] = df_gtc_raw['Volume'].apply(clean_volume)
    df_gtc_raw['GTC_Pct_Clean'] = df_gtc_raw['% GTC'].apply(clean_pct)
    df_gtc_raw['Gan_Pct_Clean'] = df_gtc_raw['% Gán'].apply(clean_pct)
    df_gtc_raw['GTC_Vol_Clean'] = df_gtc_raw['Vol_Clean'] * df_gtc_raw['GTC_Pct_Clean']
    df_gtc_raw['Gan_Vol_Clean'] = df_gtc_raw['Vol_Clean'] * df_gtc_raw['Gan_Pct_Clean']
    
    # Map AM and Province info
    ams = []
    provinces = []
    for idx, row in df_gtc_raw.iterrows():
        am_name, prov_name = resolve_po_info(row['Chi tiết'], cocau_map)
        ams.append(am_name)
        provinces.append(prov_name)
    df_gtc_raw['AM'] = ams
    df_gtc_raw['Tỉnh'] = provinces
    
    # Filter N-1 and N-2 datasets
    df_n1 = df_gtc_raw[df_gtc_raw['date_parsed'] == latest_date_str].copy()
    df_n2 = df_gtc_raw[df_gtc_raw['date_parsed'] == prev_date_str].copy() if prev_date_str else pd.DataFrame()
    
    # ══════════════════════════════════════════
    # BẢNG 1: TOP 20 BƯU CỤC TỆ NHẤT GTC CA 1 TTS (N-1)
    # ══════════════════════════════════════════
    # Group by post office
    po_n1 = df_n1.groupby(['Chi tiết', 'AM', 'Tỉnh']).agg(
        n1_vol=('Vol_Clean', 'sum'),
        n1_gtc=('GTC_Vol_Clean', 'sum'),
        n1_gan=('Gan_Vol_Clean', 'sum')
    ).reset_index()
    
    po_n1['n1_rate'] = (po_n1['n1_gtc'] / po_n1['n1_vol'] * 100).fillna(0.0)
    po_n1['n1_gan_rate'] = (po_n1['n1_gan'] / po_n1['n1_vol'] * 100).fillna(0.0)
    po_n1['san_luong'] = po_n1['n1_gtc']  # Sản lượng = Volume * %GTC
    
    # Exclude 0 volume and sort ascending by rate
    worst_pos = po_n1[po_n1['n1_vol'] > 0].sort_values(by=['n1_rate', 'n1_vol'], ascending=[True, False]).head(20).reset_index(drop=True)
    
    t1_rows_html = ""
    for idx, row in worst_pos.iterrows():
        po_name = row['Chi tiết']
        am_name = row['AM']
        tinh_name = row['Tỉnh']
        vol = row['n1_vol']
        gtc = row['n1_gtc']
        rate = row['n1_rate']
        gan_rate = row['n1_gan_rate']
        san_luong = row['san_luong']
        
        badge_class = "rate-green" if rate >= 92.0 else "rate-yellow" if rate >= 85.0 else "rate-red"
        dot_class = "dot-green" if rate >= 92.0 else "dot-yellow" if rate >= 85.0 else "dot-red"
        gan_badge = "rate-green" if gan_rate >= 95.0 else "rate-yellow" if gan_rate >= 85.0 else "rate-red"
        gan_dot = "dot-green" if gan_rate >= 95.0 else "dot-yellow" if gan_rate >= 85.0 else "dot-red"
        
        t1_rows_html += f"""
        <tr>
            <td class="number">{idx + 1}</td>
            <td class="am-name">{po_name}</td>
            <td style="font-weight:500;">{am_name}</td>
            <td>{tinh_name}</td>
            <td class="number">{vol:,.0f}</td>
            <td class="number">{san_luong:,.0f}</td>
            <td style="text-align:center;">
                <span class="rate-badge {gan_badge}"><span class="dot {gan_dot}"></span>{gan_rate:.2f}%</span>
            </td>
            <td style="text-align:center;">
                <span class="rate-badge {badge_class}"><span class="dot {dot_class}"></span>{rate:.2f}%</span>
            </td>
        </tr>
        """
        
    # ══════════════════════════════════════════
    # BẢNG 2: TỶ LỆ GTC CA 1 TTS THEO AM (WORST TO BEST)
    # ══════════════════════════════════════════
    # Summarize N-1 by AM
    am_n1 = df_n1.groupby('AM').agg(
        n1_vol=('Vol_Clean', 'sum'),
        n1_gtc=('GTC_Vol_Clean', 'sum'),
        n1_gan=('Gan_Vol_Clean', 'sum')
    ).reset_index()
    am_n1['n1_rate'] = (am_n1['n1_gtc'] / am_n1['n1_vol'] * 100).fillna(100.0)
    am_n1['n1_gan_rate'] = (am_n1['n1_gan'] / am_n1['n1_vol'] * 100).fillna(0.0)
    am_n1['san_luong'] = am_n1['n1_gtc']  # Sản lượng = Volume * %GTC
    
    # Summarize N-2 by AM
    am_n2_map = {}
    if not df_n2.empty:
        am_n2 = df_n2.groupby('AM').agg(
            n2_vol=('Vol_Clean', 'sum'),
            n2_gtc=('GTC_Vol_Clean', 'sum')
        ).reset_index()
        am_n2['n2_rate'] = (am_n2['n2_gtc'] / am_n2['n2_vol'] * 100).fillna(100.0)
        am_n2_map = dict(zip(am_n2['AM'], am_n2['n2_rate']))
        
    # Sort worst first (rate ascending)
    am_n1 = am_n1.sort_values(by='n1_rate', ascending=True).reset_index(drop=True)
    
    t2_rows_html = ""
    for idx, row in am_n1.iterrows():
        am_name = row['AM']
        vol = row['n1_vol']
        gtc = row['n1_gtc']
        rate = row['n1_rate']
        gan_rate = row['n1_gan_rate']
        san_luong = row['san_luong']
        
        badge_class = "rate-green" if rate >= 92.0 else "rate-yellow" if rate >= 85.0 else "rate-red"
        dot_class = "dot-green" if rate >= 92.0 else "dot-yellow" if rate >= 85.0 else "dot-red"
        gan_badge = "rate-green" if gan_rate >= 95.0 else "rate-yellow" if gan_rate >= 85.0 else "rate-red"
        gan_dot = "dot-green" if gan_rate >= 95.0 else "dot-yellow" if gan_rate >= 85.0 else "dot-red"
        
        # Day over day difference
        if prev_date_str and am_name in am_n2_map:
            prev_rate = am_n2_map[am_name]
            diff = rate - prev_rate
            if abs(diff) < 0.01:
                diff_html = '<span class="compare-same">—</span>'
            elif diff > 0:
                diff_html = f'<span class="compare-up">▲ +{diff:.2f}%</span>'
            else:
                diff_html = f'<span class="compare-down">▼ {diff:.2f}%</span>'
        else:
            diff_html = '<span class="compare-same">N/A</span>'
            
        t2_rows_html += f"""
        <tr>
            <td class="number">{idx + 1}</td>
            <td class="am-name">{am_name}</td>
            <td class="number">{vol:,.0f}</td>
            <td class="number">{san_luong:,.0f}</td>
            <td style="text-align:center;">
                <span class="rate-badge {gan_badge}"><span class="dot {gan_dot}"></span>{gan_rate:.2f}%</span>
            </td>
            <td style="text-align:center;">
                <span class="rate-badge {badge_class}"><span class="dot {dot_class}"></span>{rate:.2f}%</span>
            </td>
            <td style="text-align:center;">{diff_html}</td>
        </tr>
        """
        
    # Region Total for Table 2
    total_vol_n1 = am_n1['n1_vol'].sum()
    total_gtc_n1 = am_n1['n1_gtc'].sum()
    region_rate_n1 = (total_gtc_n1 / total_vol_n1 * 100) if total_vol_n1 > 0 else 100.0
    
    if prev_date_str and not df_n2.empty:
        total_vol_n2 = df_n2['Vol_Clean'].sum()
        total_gtc_n2 = df_n2['GTC_Vol_Clean'].sum()
        region_rate_n2 = (total_gtc_n2 / total_vol_n2 * 100) if total_vol_n2 > 0 else 100.0
        reg_diff = region_rate_n1 - region_rate_n2
        if abs(reg_diff) < 0.01:
            reg_diff_html = '<span class="compare-same">—</span>'
            reg_diff_txt = "không đổi"
        elif reg_diff > 0:
            reg_diff_html = f'<span class="compare-up">▲ +{reg_diff:.2f}%</span>'
            reg_diff_txt = f"tăng +{reg_diff:.2f}%"
        else:
            reg_diff_html = f'<span class="compare-down">▼ {reg_diff:.2f}%</span>'
            reg_diff_txt = f"giảm {reg_diff:.2f}%"
        reg_diff_msg = f"(<b>{reg_diff_txt}</b> so với hôm qua)"
    else:
        reg_diff_html = '<span class="compare-same">N/A</span>'
        reg_diff_msg = "(so với hôm qua: N/A)"
        
    region_badge = "rate-green" if region_rate_n1 >= 92.0 else "rate-yellow" if region_rate_n1 >= 85.0 else "rate-red"
    region_dot = "dot-green" if region_rate_n1 >= 92.0 else "dot-yellow" if region_rate_n1 >= 85.0 else "dot-red"
    
    total_gan_n1 = am_n1['n1_gan'].sum()
    region_gan_rate = (total_gan_n1 / total_vol_n1 * 100) if total_vol_n1 > 0 else 0.0
    region_gan_badge = "rate-green" if region_gan_rate >= 95.0 else "rate-yellow" if region_gan_rate >= 85.0 else "rate-red"
    region_gan_dot = "dot-green" if region_gan_rate >= 95.0 else "dot-yellow" if region_gan_rate >= 85.0 else "dot-red"
    
    t2_rows_html += f"""
    <tr class="total-row">
        <td colspan="2">TOÀN VÙNG NTB</td>
        <td class="number">{total_vol_n1:,.0f}</td>
        <td class="number">{total_gtc_n1:,.0f}</td>
        <td style="text-align:center;">
            <span class="rate-badge {region_gan_badge}"><span class="dot {region_gan_dot}"></span>{region_gan_rate:.2f}%</span>
        </td>
        <td style="text-align:center;">
            <span class="rate-badge {region_badge}"><span class="dot {region_dot}"></span>{region_rate_n1:.2f}%</span>
        </td>
        <td style="text-align:center;">{reg_diff_html}</td>
    </tr>
    """
    
    # ══════════════════════════════════════════
    # BẢNG 3: VIEW TỔNG QUAN GTC CA 1 TTS THEO NGÀY
    # ══════════════════════════════════════════
    # Group by date for all available dates
    daily_stats = df_gtc_raw.groupby('date_parsed').agg(
        vol=('Vol_Clean', 'sum'),
        gtc=('GTC_Vol_Clean', 'sum'),
        gan=('Gan_Vol_Clean', 'sum')
    ).reset_index()
    daily_stats['rate'] = (daily_stats['gtc'] / daily_stats['vol'] * 100).fillna(100.0)
    
    # Get last 7 days sorted chronologically
    daily_stats = daily_stats.sort_values(by='date_parsed').tail(7).reset_index(drop=True)
    
    t3_rows_html = ""
    for idx, row in daily_stats.iterrows():
        d_str = row['date_parsed']
        d_obj = datetime.strptime(d_str, "%Y-%m-%d")
        d_display = d_obj.strftime("%d/%m/%Y")
        d_vol = row['vol']
        d_gtc = row['gtc']
        d_rate = row['rate']
        d_gan = row['gan']
        d_gan_rate = (d_gan / d_vol * 100) if d_vol > 0 else 0.0
        
        badge_class = "rate-green" if d_rate >= 92.0 else "rate-yellow" if d_rate >= 85.0 else "rate-red"
        dot_class = "dot-green" if d_rate >= 92.0 else "dot-yellow" if d_rate >= 85.0 else "dot-red"
        gan_badge = "rate-green" if d_gan_rate >= 95.0 else "rate-yellow" if d_gan_rate >= 85.0 else "rate-red"
        gan_dot = "dot-green" if d_gan_rate >= 95.0 else "dot-yellow" if d_gan_rate >= 85.0 else "dot-red"
        
        # Day over day compare
        if idx > 0:
            prev_day_rate = daily_stats.loc[idx - 1, 'rate']
            d_diff = d_rate - prev_day_rate
            if abs(d_diff) < 0.01:
                d_diff_html = '<span class="compare-same">—</span>'
            elif d_diff > 0:
                d_diff_html = f'<span class="compare-up">▲ +{d_diff:.2f}%</span>'
            else:
                d_diff_html = f'<span class="compare-down">▼ {d_diff:.2f}%</span>'
        else:
            d_diff_html = '<span class="compare-same">N/A</span>'
            
        t3_rows_html += f"""
        <tr>
            <td style="font-weight:600; text-align:center;">{d_display}</td>
            <td class="number">{d_vol:,.0f}</td>
            <td class="number">{d_gtc:,.0f}</td>
            <td style="text-align:center;">
                <span class="rate-badge {gan_badge}"><span class="dot {gan_dot}"></span>{d_gan_rate:.2f}%</span>
            </td>
            <td style="text-align:center;">
                <span class="rate-badge {badge_class}"><span class="dot {dot_class}"></span>{d_rate:.2f}%</span>
            </td>
            <td style="text-align:center;">{d_diff_html}</td>
        </tr>
        """
        
    # ══════════════════════════════════════════
    # HTML TEMPLATES FOR IMAGES
    # ══════════════════════════════════════════
    # Base CSS template style
    base_css = """
        body {
            font-family: 'Inter', sans-serif;
            background-color: #eef2ff;
            margin: 0;
            padding: 20px;
            display: flex;
            justify-content: center;
            align-items: center;
        }
        #capture-container {
            background: #ffffff;
            padding: 28px;
            border-radius: 16px;
            box-shadow: 0 4px 24px rgba(67, 56, 202, 0.08);
            border: 1px solid #e0e7ff;
            width: 900px;
        }
        .header {
            margin-bottom: 20px;
            text-align: center;
            background: linear-gradient(135deg, #4338ca 0%, #6366f1 100%);
            padding: 16px 20px;
            border-radius: 12px;
        }
        .header h2 {
            margin: 0;
            font-family: 'Plus Jakarta Sans', sans-serif;
            font-size: 19px;
            color: #ffffff;
            font-weight: 800;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        .header p {
            margin: 4px 0 0 0;
            font-size: 13px;
            color: #c7d2fe;
            font-weight: 500;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            text-align: left;
        }
        th {
            font-family: 'Plus Jakarta Sans', sans-serif;
            background-color: #312e81;
            color: #e0e7ff;
            font-weight: 700;
            font-size: 11px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            padding: 10px 8px;
            border-bottom: 2px solid #4338ca;
        }
        td {
            padding: 8px 8px;
            font-size: 12px;
            border-bottom: 1px solid #f1f5f9;
            color: #334155;
        }
        tr:nth-child(even) {
            background-color: #f8faff;
        }
        tr:hover {
            background-color: #eef2ff;
        }
        .number {
            text-align: right;
            font-weight: 500;
        }
        .am-name {
            font-weight: 600;
            color: #1e1b4b;
        }
        .rate-badge {
            display: inline-flex;
            align-items: center;
            padding: 4px 10px;
            border-radius: 9999px;
            font-weight: 700;
            font-size: 12px;
            line-height: 1;
        }
        .rate-green {
            background-color: #dcfce7;
            color: #15803d;
            border: 1px solid #bbf7d0;
        }
        .rate-yellow {
            background-color: #fef9c3;
            color: #a16207;
            border: 1px solid #fde68a;
        }
        .rate-red {
            background-color: #fee2e2;
            color: #dc2626;
            border: 1px solid #fecaca;
        }
        .dot {
            height: 6px;
            width: 6px;
            border-radius: 50%;
            display: inline-block;
            margin-right: 5px;
        }
        .dot-green { background-color: #16a34a; }
        .dot-yellow { background-color: #ca8a04; }
        .dot-red { background-color: #dc2626; }
        .compare-up {
            color: #16a34a;
            font-weight: 600;
            font-size: 11px;
        }
        .compare-down {
            color: #dc2626;
            font-weight: 600;
            font-size: 11px;
        }
        .compare-same {
            color: #94a3b8;
            font-weight: 500;
        }
        .total-row {
            background-color: #eef2ff;
            border-top: 2px solid #4338ca;
        }
        .total-row td {
            font-weight: 700;
            color: #1e1b4b;
            font-size: 12px;
            border-bottom: 2px solid #4338ca;
        }
    """
    
    t1_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="utf-8">
    <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <style>{base_css}</style>
    </head>
    <body>
    <div id="capture-container">
        <div class="header">
            <h2>Bảng 1: Top 20 Bưu Cục Tệ Nhất GTC Ca 1 TTS</h2>
            <p>Khu vực NTB — Báo cáo ngày {latest_date_display}</p>
        </div>
        <table>
            <thead>
                <tr>
                    <th style="width: 5%; text-align:center;">STT</th>
                    <th style="width: 25%;">Bưu cục</th>
                    <th style="width: 18%;">AM Quản lý</th>
                    <th style="width: 10%;">Tỉnh</th>
                    <th style="width: 10%; text-align:right;">Volume</th>
                    <th style="width: 10%; text-align:right;">Sản lượng</th>
                    <th style="width: 11%; text-align:center;">% Gán</th>
                    <th style="width: 11%; text-align:center;">Tỷ lệ GTC ca 1</th>
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
    
    t2_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="utf-8">
    <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <style>{base_css}</style>
    </head>
    <body>
    <div id="capture-container">
        <div class="header">
            <h2>Bảng 2: Tỷ Lệ GTC Ca 1 TTS Theo AM</h2>
            <p>Khu vực NTB — Báo cáo ngày {latest_date_display} (Sắp xếp từ tệ đến tốt)</p>
        </div>
        <table>
            <thead>
                <tr>
                    <th style="width: 7%; text-align:center;">STT</th>
                    <th style="width: 25%;">AM Phụ trách</th>
                    <th style="width: 12%; text-align:right;">Volume</th>
                    <th style="width: 12%; text-align:right;">Sản lượng</th>
                    <th style="width: 12%; text-align:center;">% Gán</th>
                    <th style="width: 12%; text-align:center;">Tỷ lệ GTC ca 1</th>
                    <th style="width: 10%; text-align:center;">So ngày cũ</th>
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
    
    t3_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="utf-8">
    <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <style>{base_css}</style>
    </head>
    <body>
    <div id="capture-container">
        <div class="header">
            <h2>Bảng 3: Tổng Quan Chỉ Số GTC Ca 1 TTS Toàn Vùng Theo Ngày</h2>
            <p>Khu vực NTB — Xu hướng 7 ngày gần nhất</p>
        </div>
        <table>
            <thead>
                <tr>
                    <th style="width: 18%; text-align:center;">Ngày báo cáo</th>
                    <th style="width: 15%; text-align:right;">Tổng Volume</th>
                    <th style="width: 15%; text-align:right;">Sản lượng</th>
                    <th style="width: 14%; text-align:center;">% Gán</th>
                    <th style="width: 14%; text-align:center;">Tỷ lệ GTC ca 1</th>
                    <th style="width: 12%; text-align:center;">Biến động</th>
                </tr>
            </thead>
            <tbody>
                {t3_rows_html}
            </tbody>
        </table>
    </div>
    </body>
    </html>
    """
    
    # 5. Write HTML files
    t1_html_path = os.path.join(BASE_DIR, "t1_gtc_worst_pos.html")
    t2_html_path = os.path.join(BASE_DIR, "t2_gtc_am_rates.html")
    t3_html_path = os.path.join(BASE_DIR, "t3_gtc_daily_overview.html")
    
    with open(t1_html_path, "w", encoding="utf-8") as f:
        f.write(t1_html)
    with open(t2_html_path, "w", encoding="utf-8") as f:
        f.write(t2_html)
    with open(t3_html_path, "w", encoding="utf-8") as f:
        f.write(t3_html)
        
    # 6. Take screenshots using Playwright
    t1_image_path = os.path.join(BASE_DIR, "gtc_worst_pos.png")
    t2_image_path = os.path.join(BASE_DIR, "gtc_am_rates.png")
    t3_image_path = os.path.join(BASE_DIR, "gtc_daily_overview.png")
    
    print("📸 Đang dựng và chụp ảnh các bảng báo cáo bằng Playwright...")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(viewport={"width": 1000, "height": 1200})
        
        # Screenshot Table 1
        page.goto(f"file:///{t1_html_path.replace('\\', '/')}")
        page.wait_for_timeout(1000)
        page.locator("#capture-container").screenshot(path=t1_image_path)
        print(f"✔️ Đã xuất ảnh bảng 1: {t1_image_path}")
        
        # Screenshot Table 2
        page.goto(f"file:///{t2_html_path.replace('\\', '/')}")
        page.wait_for_timeout(1000)
        page.locator("#capture-container").screenshot(path=t2_image_path)
        print(f"✔️ Đã xuất ảnh bảng 2: {t2_image_path}")
        
        # Screenshot Table 3
        page.goto(f"file:///{t3_html_path.replace('\\', '/')}")
        page.wait_for_timeout(1000)
        page.locator("#capture-container").screenshot(path=t3_image_path)
        print(f"✔️ Đã xuất ảnh bảng 3: {t3_image_path}")
        
        browser.close()
        
    # Cleanup temp HTML files
    try:
        os.remove(t1_html_path)
        os.remove(t2_html_path)
        os.remove(t3_html_path)
    except:
        pass
        
    # Copy screenshots to artifact folder for user viewing if directory exists
    artifact_dir = r"C:\Users\lap4all\.gemini\antigravity-ide\brain\d705cfae-7bfd-453c-b738-07a6c2618aea"
    if os.path.exists(artifact_dir):
        import shutil
        try:
            shutil.copy2(t1_image_path, os.path.join(artifact_dir, "gtc_worst_pos.png"))
            shutil.copy2(t2_image_path, os.path.join(artifact_dir, "gtc_am_rates.png"))
            shutil.copy2(t3_image_path, os.path.join(artifact_dir, "gtc_daily_overview.png"))
            print("📋 Đã copy các file ảnh báo cáo vào thư mục brain artifact.")
        except Exception as e:
            print(f"⚠️ Cảnh báo copy ảnh báo cáo: {e}")
            
    # 7. Post report text and screenshots to GTalk group
    print("📡 Đang phát sóng tin nhắn báo cáo lên GTalk group...")
    
    caption = f"""<b>BÁO CÁO TỶ LỆ GTC CA 1 TTS NTB NGÀY {latest_date_display}</b>

<b>1. Tỷ lệ GTC Ca 1 TTS toàn vùng NTB:</b> <b>{region_rate_n1:.2f}%</b> {reg_diff_msg}
- Giao thành công: <b>{total_gtc_n1:,.0f} đơn</b>
- Tổng Volume ca 1: <b>{total_vol_n1:,.0f} đơn</b>

<b>2. Top 3 AM tệ nhất GTC ca 1 TTS:</b>"""

    # Add top 3 worst AMs in text summary
    for idx, row in am_n1.head(3).iterrows():
        am_name = row['AM']
        rate = row['n1_rate']
        vol = row['n1_vol']
        caption += f"\n- AM <b>{am_name}</b>: <b>{rate:.2f}%</b> (Volume {vol:,.0f})"

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
            "Metadata": json.dumps({"width": 850, "height": 1000}),
            "oaToken": GTALK_OA_TOKEN
        }
        
        print(f"📡  [GTalk Upload] Đang khởi tạo upload cho {file_name} (Size: {file_size} bytes)...")
        resp_init = requests.post("https://mbff.ghn.vn/api/gtalk/initiate-upload", json=init_payload, timeout=20)
        if resp_init.status_code != 200:
            print(f"❌  [GTalk Upload] Lỗi khởi tạo upload, status={resp_init.status_code}, response={resp_init.text}")
            return None
            
        init_data = resp_init.json()
        if init_data.get("errorCode") != "success":
            print(f"❌  [GTalk Upload] API báo lỗi khi khởi tạo: {init_data}")
            return None
            
        presigned_url = init_data["data"]["PresignedURL"]
        upload_id = init_data["data"]["UploadId"]
        
        print(f"📡  [GTalk Upload] Đang PUT dữ liệu ảnh lên S3 presigned URL...")
        resp_put = requests.put(presigned_url, data=file_bytes, headers={"Content-Type": "image/png"}, timeout=30)
        if resp_put.status_code != 200:
            print(f"❌  [GTalk Upload] Lỗi PUT ảnh lên S3, status={resp_put.status_code}, response={resp_put.text}")
            return None
            
        print(f"📡  [GTalk Upload] Đang xác nhận hoàn tất upload (UploadId: {upload_id})...")
        resp_comp = requests.post("https://mbff.ghn.vn/api/gtalk/complete-upload", json={"oaToken": GTALK_OA_TOKEN, "UploadId": upload_id}, timeout=20)
        if resp_comp.status_code != 200:
            print(f"❌  [GTalk Upload] Lỗi hoàn tất upload, status={resp_comp.status_code}, response={resp_comp.text}")
            return None
            
        comp_data = resp_comp.json()
        if comp_data.get("errorCode") != "success":
            print(f"❌  [GTalk Upload] API báo lỗi khi hoàn tất: {comp_data}")
            return None
            
        file_id = comp_data["data"]["Id"]
        print(f"✅  [GTalk Upload] Upload thành công! File ID: {file_id}")
        return file_id

    # Upload all 3 images
    img1_id = upload_image_to_gtalk(t1_image_path)
    img2_id = upload_image_to_gtalk(t2_image_path)
    img3_id = upload_image_to_gtalk(t3_image_path)
    
    if img1_id and img2_id and img3_id:
        send_payload = {
            "channelId": GTALK_CHANNEL_ID,
            "clientMsgId": str(int(datetime.now().timestamp() * 1000)),
            "content": {
                "parseMode": "HTML",
                "attachment": {
                    "caption": caption,
                    "items": [
                        {"image": {"fileId": img1_id, "width": 850, "height": 1000}},
                        {"image": {"fileId": img2_id, "width": 850, "height": 1000}},
                        {"image": {"fileId": img3_id, "width": 850, "height": 1000}}
                    ]
                }
            },
            "oaToken": GTALK_OA_TOKEN
        }
        r_send = requests.post("https://mbff.ghn.vn/api/gtalk/send-message", json=send_payload, timeout=25)
        if r_send.status_code == 200 and r_send.json().get("errorCode") == "success":
            print("✅ Đã gửi báo cáo và ảnh bảng biểu lên GTalk thành công!")
        else:
            print(f"❌ Lỗi gửi tin nhắn GTalk: {r_send.text}")
    else:
        print(f"❌ Lỗi tải ảnh lên GTalk. img1_id: {img1_id}, img2_id: {img2_id}, img3_id: {img3_id}")

    # Cleanup local screenshots
    try:
        os.remove(t1_image_path)
        os.remove(t2_image_path)
        os.remove(t3_image_path)
    except:
        pass
        
    print("🎉 BÁO CÁO GTC CA 1 TTS ĐÃ HOÀN THÀNH!")

if __name__ == "__main__":
    main()
