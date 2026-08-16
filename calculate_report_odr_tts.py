import sys
import io
import os
import re
import json
import pandas as pd
import gspread
import unicodedata
import requests
from datetime import datetime, timedelta
from google.oauth2.service_account import Credentials
from google.oauth2.credentials import Credentials as UserCredentials
from playwright.sync_api import sync_playwright
from dotenv import load_dotenv

# Fix encoding for console execution/logging
os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace', write_through=True)
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace', write_through=True)
except AttributeError:
    pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
SHEET_KEY = '1705_0rKkgXBpsCbgK10EDr_mzSGhJOAcCa1WZsrWrU4'
GTALK_OA_TOKEN = os.environ.get("ODR_TTS_GTALK_OA_TOKEN") or "2067164759497973760:RfgqBJY4QtV18udu1wMEfhmoRNI4hgBv"
GTALK_CHANNEL_ID = os.environ.get("ODR_TTS_GTALK_CHANNEL_ID") or "2073028212579307520"

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

SERVICE_ACCOUNT_CANDIDATES = [
    os.path.join(BASE_DIR, 'credentials.json'),
    r'C:\Users\lap4all\Documents\Auto report\credentials.json',
    r'C:\Users\lap4all\Desktop\Backlog_Automation\credentials.json',
    'credentials.json'
]

AUTH_USER_CANDIDATES = [
    os.path.join(BASE_DIR, 'authorized_user.json'),
    r'C:\Users\lap4all\Documents\Auto report\authorized_user.json',
    r'C:\Users\lap4all\Desktop\Backlog_Automation\authorized_user.json',
    'authorized_user.json'
]


def get_gspread_client(spreadsheet_id=None):
    """Thử Service Account trước, fallback sang authorized_user.json nếu bị từ chối quyền."""
    for cred_path in SERVICE_ACCOUNT_CANDIDATES:
        if os.path.isfile(cred_path):
            try:
                creds = Credentials.from_service_account_file(cred_path, scopes=SCOPES)
                gc = gspread.authorize(creds)
                if spreadsheet_id:
                    gc.open_by_key(spreadsheet_id)
                return gc
            except Exception as e:
                print(f"⚠️ Service account ({cred_path}) không có quyền: {e}. Đang chuyển sang authorized_user.json...", flush=True)

    for auth_file in AUTH_USER_CANDIDATES:
        if os.path.exists(auth_file):
            try:
                creds = UserCredentials.from_authorized_user_file(auth_file, scopes=SCOPES)
                gc = gspread.authorize(creds)
                if spreadsheet_id:
                    gc.open_by_key(spreadsheet_id)
                return gc
            except Exception:
                pass

    raise PermissionError("Không thể xác thực Google Sheets bằng credentials.json hoặc authorized_user.json")


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

def main():
    print("🚀 BẮT ĐẦU CHẠY BÁO CÁO ODR & ODR TTS LÚC:", datetime.now().strftime('%H:%M:%S'))
    
    # 1. Authorize Google Sheet
    print("📡 Connecting to Google Sheets...")
    gc_client = get_gspread_client(spreadsheet_id=SHEET_KEY)
    sh = gc_client.open_by_key(SHEET_KEY)
    
    # 2. Read Worksheets
    print("📥 Loading data worksheets...")
    ws_cc = None
    for sname in ["CoCauVung", "Cocau", "Cơ cấu", "cơ cấu"]:
        try:
            ws_cc = sh.worksheet(sname)
            break
        except Exception:
            pass
    if not ws_cc:
        print("❌ Không tìm thấy tab CoCauVung hoặc Cơ cấu trong spreadsheet.")
        sys.exit(1)
        
    df_cocau = pd.DataFrame(ws_cc.get_all_records())
    
    try:
        df_odr = pd.DataFrame(sh.worksheet("ODR").get_all_records())
        if df_odr.empty or 'Time' not in df_odr.columns:
            raise ValueError("Tab 'ODR' rỗng hoặc thiếu cột 'Time'")
    except Exception as e:
        print(f"⚠️ Tab 'ODR' trên Sheet chính bị lỗi IMPORTRANGE ({e}). Tự động đọc dữ liệu trực tiếp từ Master Sheet 1DAwY-46twFrHIs77R4p4IMuIZ6JTE-e58Aj-9Kcr5Jk...")
        try:
            sh_master = get_gspread_client('1DAwY-46twFrHIs77R4p4IMuIZ6JTE-e58Aj-9Kcr5Jk').open_by_key('1DAwY-46twFrHIs77R4p4IMuIZ6JTE-e58Aj-9Kcr5Jk')
            df_odr = pd.DataFrame(sh_master.worksheet("ODR TTS").get_all_records())
        except Exception as e2:
            print(f"❌ Không thể đọc dữ liệu ODR từ Master Sheet: {e2}")
            sys.exit(1)

    df_tts = pd.DataFrame(sh.worksheet("ODR - TTS").get_all_records())
    
    # 3. Build Cơ cấu map
    cocau_map = {}
    for idx, row in df_cocau.iterrows():
        bc_name = normalize_str(row['Bưu cục']).lower()
        cocau_map[bc_name] = (row['AM'], row['Tỉnh'])
        
    # 4. Parse Dates and Get target_date (N-1) and comparison date (N-2)
    # Format of Time column is e.g. "2026-06-27 - Thứ 7"
    df_odr['date_parsed'] = df_odr['Time'].apply(lambda x: x.split(" - ")[0].strip() if " - " in str(x) else str(x).strip())
    unique_dates = sorted(df_odr['date_parsed'].unique())
    if not unique_dates:
        print("❌ No dates found in ODR sheet.")
        sys.exit(1)
        
    latest_date_str = unique_dates[-1]
    prev_date_str = unique_dates[-2] if len(unique_dates) > 1 else None
    
    # Format dates for display
    latest_date_obj = datetime.strptime(latest_date_str, "%Y-%m-%d")
    latest_date_display = latest_date_obj.strftime("%d/%m/%Y")
    
    print(f"📅 Target Date (N-1): {latest_date_str} (Display: {latest_date_display})")
    if prev_date_str:
        print(f"📅 Comparison Date (N-2): {prev_date_str}")
    else:
        print("📅 Comparison Date (N-2): N/A")
        
    # Helper to parse %Ontime string to float
    def parse_percent(val):
        if not val:
            return 1.0
        val_str = str(val).replace(",", ".").replace("%", "").strip()
        try:
            return float(val_str) / 100
        except ValueError:
            return 1.0
            
    # --- TABLE 1 CALCULATIONS: ODR Rate by AM ---
    # We filter ODR rows for N-1 and N-2
    df_odr_n1 = df_odr[df_odr['date_parsed'] == latest_date_str].copy()
    df_odr_n2 = df_odr[df_odr['date_parsed'] == prev_date_str].copy() if prev_date_str else pd.DataFrame()
    
    # Map AM and Province
    for df in [df_odr_n1, df_odr_n2]:
        if not df.empty:
            ams = []
            provinces = []
            for idx, row in df.iterrows():
                po_name = row['Chi tiết']
                am, prov = resolve_po_info(po_name, cocau_map)
                ams.append(am)
                provinces.append(prov)
            df['AM'] = ams
            df['Tỉnh'] = provinces
            df['GTC'] = pd.to_numeric(df['GTC'], errors='coerce').fillna(0)
            df['Ontime_rate'] = df['%Ontime'].apply(parse_percent)
            df['Ontime_volume'] = df['GTC'] * df['Ontime_rate']
            
    # Summarize N-1 by AM
    am_odr_n1 = df_odr_n1.groupby('AM').agg(
        n1_can=('GTC', 'sum'),
        n1_ontime=('Ontime_volume', 'sum')
    ).reset_index()
    am_odr_n1['n1_rate'] = (am_odr_n1['n1_ontime'] / am_odr_n1['n1_can'] * 100).fillna(100.0)
    
    # Summarize N-2 by AM if available
    am_odr_n2_map = {}
    if not df_odr_n2.empty:
        am_odr_n2 = df_odr_n2.groupby('AM').agg(
            n2_can=('GTC', 'sum'),
            n2_ontime=('Ontime_volume', 'sum')
        ).reset_index()
        am_odr_n2['n2_rate'] = (am_odr_n2['n2_ontime'] / am_odr_n2['n2_can'] * 100).fillna(100.0)
        am_odr_n2_map = dict(zip(am_odr_n2['AM'], am_odr_n2['n2_rate']))
        
    # Sort N-1 ODR rate descending
    am_odr_n1 = am_odr_n1.sort_values(by='n1_rate', ascending=False).reset_index(drop=True)
    
    # Build Table 1 Rows HTML
    t1_rows_html = ""
    for idx, row in am_odr_n1.iterrows():
        am_name = row['AM']
        vol_can = row['n1_can']
        vol_ontime = row['n1_ontime']
        rate = row['n1_rate']
        
        # Color Badge Class
        if rate >= 95.0:
            badge_class = "rate-green"
            dot_class = "dot-green"
        elif rate >= 90.0:
            badge_class = "rate-yellow"
            dot_class = "dot-yellow"
        else:
            badge_class = "rate-red"
            dot_class = "dot-red"
            
        # Day-over-day difference comparison
        if prev_date_str and am_name in am_odr_n2_map:
            prev_rate = am_odr_n2_map[am_name]
            diff = rate - prev_rate
            if abs(diff) < 0.01:
                diff_html = '<span class="compare-same">—</span>'
            elif diff > 0:
                diff_html = f'<span class="compare-up">▲ +{diff:.2f}%</span>'
            else:
                diff_html = f'<span class="compare-down">▼ {diff:.2f}%</span>'
        else:
            diff_html = '<span class="compare-same">N/A</span>'
            
        t1_rows_html += f"""
        <tr>
            <td class="number">{idx + 1}</td>
            <td class="am-name">{am_name}</td>
            <td class="number">{vol_can:,.0f}</td>
            <td class="number">{vol_ontime:,.0f}</td>
            <td style="text-align:center;">
                <span class="rate-badge {badge_class}"><span class="dot {dot_class}"></span>{rate:.2f}%</span>
            </td>
            <td style="text-align:center;">{diff_html}</td>
        </tr>
        """
        
    # Grand Total for Table 1
    total_n1_can = am_odr_n1['n1_can'].sum()
    total_n1_ontime = am_odr_n1['n1_ontime'].sum()
    grand_odr_rate = (total_n1_ontime / total_n1_can * 100) if total_n1_can > 0 else 100.0
    
    # Grand total comparison diff
    if prev_date_str and not df_odr_n2.empty:
        total_n2_can = df_odr_n2['GTC'].sum()
        total_n2_ontime = df_odr_n2['Ontime_volume'].sum()
        grand_odr_rate_n2 = (total_n2_ontime / total_n2_can * 100) if total_n2_can > 0 else 100.0
        tot_diff = grand_odr_rate - grand_odr_rate_n2
        if abs(tot_diff) < 0.01:
            tot_diff_html = '<span class="compare-same">—</span>'
            tot_diff_txt = "không đổi"
        elif tot_diff > 0:
            tot_diff_html = f'<span class="compare-up">▲ +{tot_diff:.2f}%</span>'
            tot_diff_txt = f"tăng +{tot_diff:.2f}%"
        else:
            tot_diff_html = f'<span class="compare-down">▼ {tot_diff:.2f}%</span>'
            tot_diff_txt = f"giảm {tot_diff:.2f}%"
        tot_diff_msg = f"(<b>{tot_diff_txt}</b> so với hôm qua)"
    else:
        tot_diff_html = '<span class="compare-same">N/A</span>'
        tot_diff_msg = "(so với hôm qua: N/A)"
        
    grand_badge = "rate-green" if grand_odr_rate >= 95.0 else "rate-yellow" if grand_odr_rate >= 90.0 else "rate-red"
    grand_dot = "dot-green" if grand_odr_rate >= 95.0 else "dot-yellow" if grand_odr_rate >= 90.0 else "dot-red"
    
    t1_rows_html += f"""
    <tr class="total-row">
        <td colspan="2">TỔNG CỘNG NTB</td>
        <td class="number">{total_n1_can:,.0f}</td>
        <td class="number">{total_n1_ontime:,.0f}</td>
        <td style="text-align:center;">
            <span class="rate-badge {grand_badge}"><span class="dot {grand_dot}"></span>{grand_odr_rate:.2f}%</span>
        </td>
        <td style="text-align:center;">{tot_diff_html}</td>
    </tr>
    """
    
    # --- TABLE 2 CALCULATIONS: ODR Backlog (ODR - TTS) by AM ---
    # Parse total_orders to integers
    df_tts['total_orders'] = pd.to_numeric(df_tts['total_orders'], errors='coerce').fillna(0).astype(int)
    
    # Delay categories
    delay_categories = [
        "(1) Đã trễ ODR",
        "(2) Phải giao trong hôm nay",
        "(3) Còn 2 ngày",
        "(4) Còn 3 ngày",
        "(5) Còn >3 ngày"
    ]
    
    # Pivot
    backlog_pivot = df_tts.pivot_table(
        index='AM',
        columns='nhom_tre',
        values='total_orders',
        aggfunc='sum',
        fill_value=0
    ).reset_index()
    
    # Reindex columns to make sure all delay groups exist
    for col in delay_categories:
        if col not in backlog_pivot.columns:
            backlog_pivot[col] = 0
            
    # Rearrange columns
    backlog_pivot = backlog_pivot[['AM'] + delay_categories]
    backlog_pivot['Tổng tồn'] = backlog_pivot[delay_categories].sum(axis=1)
    backlog_pivot = backlog_pivot.sort_values(by='Tổng tồn', ascending=False).reset_index(drop=True)
    
    # Build Table 2 Rows HTML
    t2_rows_html = ""
    for idx, row in backlog_pivot.iterrows():
        am_name = row['AM']
        col1 = row["(1) Đã trễ ODR"]
        col2 = row["(2) Phải giao trong hôm nay"]
        col3 = row["(3) Còn 2 ngày"]
        col4 = row["(4) Còn 3 ngày"]
        col5 = row["(5) Còn >3 ngày"]
        tot = row["Tổng tồn"]
        
        # Color coding for late/today orders
        c1_style = 'style="font-weight:700; color:#dc2626; background-color:#fee2e2;"' if col1 > 0 else ""
        c2_style = 'style="font-weight:600; color:#b45309; background-color:#fef9c3;"' if col2 > 0 else ""
        
        t2_rows_html += f"""
        <tr>
            <td class="number">{idx + 1}</td>
            <td class="am-name">{am_name}</td>
            <td class="number" {c1_style}>{col1:,.0f}</td>
            <td class="number" {c2_style}>{col2:,.0f}</td>
            <td class="number">{col3:,.0f}</td>
            <td class="number">{col4:,.0f}</td>
            <td class="number">{col5:,.0f}</td>
            <td class="number" style="font-weight:700; background-color:#f1f5f9;">{tot:,.0f}</td>
        </tr>
        """
        
    # Grand Total for Table 2
    sum_c1 = backlog_pivot["(1) Đã trễ ODR"].sum()
    sum_c2 = backlog_pivot["(2) Phải giao trong hôm nay"].sum()
    sum_c3 = backlog_pivot["(3) Còn 2 ngày"].sum()
    sum_c4 = backlog_pivot["(4) Còn 3 ngày"].sum()
    sum_c5 = backlog_pivot["(5) Còn >3 ngày"].sum()
    sum_tot = backlog_pivot["Tổng tồn"].sum()
    
    t2_rows_html += f"""
    <tr class="total-row">
        <td colspan="2">TỔNG CỘNG backlog NTB</td>
        <td class="number" style="font-weight:700; color:#dc2626; background-color:#fee2e2;">{sum_c1:,.0f}</td>
        <td class="number" style="font-weight:700; color:#b45309; background-color:#fef9c3;">{sum_c2:,.0f}</td>
        <td class="number">{sum_c3:,.0f}</td>
        <td class="number">{sum_c4:,.0f}</td>
        <td class="number">{sum_c5:,.0f}</td>
        <td class="number" style="font-weight:700; color:#1e293b; background-color:#e2e8f0;">{sum_tot:,.0f}</td>
    </tr>
    """
    
    # --- HTML TEMPLATES ---
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
            width: 800px;
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
            font-size: 22px;
            color: #ea580c;
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
            background-color: #ea580c;
            color: #ffffff;
            font-weight: 700;
            font-size: 13px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            padding: 12px 10px;
            border-bottom: 2px solid #e2e8f0;
        }}
        td {{
            padding: 10px 10px;
            font-size: 13px;
            border-bottom: 1px solid #f1f5f9;
            color: #334155;
        }}
        tr:hover {{
            background-color: #f8fafc;
        }}
        .number {{
            text-align: right;
            font-family: 'Inter', sans-serif;
            font-weight: 500;
        }}
        .am-name {{
            font-weight: 600;
            color: #0f172a;
        }}
        .rate-badge {{
            display: inline-flex;
            align-items: center;
            padding: 4px 10px;
            border-radius: 9999px;
            font-weight: 700;
            font-size: 13px;
            line-height: 1;
        }}
        .rate-green {{
            background-color: #dcfce7;
            color: #16a34a;
        }}
        .rate-yellow {{
            background-color: #fef9c3;
            color: #ca8a04;
        }}
        .rate-red {{
            background-color: #fee2e2;
            color: #dc2626;
        }}
        .dot {{
            height: 6px;
            width: 6px;
            border-radius: 50%;
            display: inline-block;
            margin-right: 6px;
        }}
        .dot-green {{ background-color: #16a34a; }}
        .dot-yellow {{ background-color: #ca8a04; }}
        .dot-red {{ background-color: #dc2626; }}
        .compare-up {{
            color: #16a34a;
            font-weight: 600;
            font-size: 12px;
        }}
        .compare-down {{
            color: #dc2626;
            font-weight: 600;
            font-size: 12px;
        }}
        .compare-same {{
            color: #94a3b8;
            font-weight: 500;
        }}
        .total-row {{
            background-color: #f8fafc;
            border-top: 2px solid #cbd5e1;
        }}
        .total-row td {{
            font-weight: 700;
            color: #0f172a;
            font-size: 13px;
            border-bottom: 2px solid #cbd5e1;
        }}
    </style>
    </head>
    <body>
    <div id="capture-container">
        <div class="header">
            <h2>Bảng xếp hạng tỷ lệ ODR TTS theo AM</h2>
            <p>NTB Region — Báo cáo ngày {latest_date_display}</p>
        </div>
        <table>
            <thead>
                <tr>
                    <th style="width: 8%; text-align:center;">STT</th>
                    <th style="width: 32%;">AM phụ trách</th>
                    <th style="width: 15%; text-align:right;">GTC thành công</th>
                    <th style="width: 15%; text-align:right;">GTC đúng hạn</th>
                    <th style="width: 15%; text-align:center;">Tỷ lệ ODR</th>
                    <th style="width: 15%; text-align:center;">So ngày cũ</th>
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
            width: 1100px;
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
            font-size: 22px;
            color: #ea580c;
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
            background-color: #ea580c;
            color: #ffffff;
            font-weight: 700;
            font-size: 13px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            padding: 12px 10px;
            border-bottom: 2px solid #e2e8f0;
        }}
        td {{
            padding: 10px 10px;
            font-size: 13px;
            border-bottom: 1px solid #f1f5f9;
            color: #334155;
        }}
        tr:hover {{
            background-color: #f8fafc;
        }}
        .number {{
            text-align: right;
            font-family: 'Inter', sans-serif;
            font-weight: 500;
        }}
        .am-name {{
            font-weight: 600;
            color: #0f172a;
        }}
        .total-row {{
            background-color: #f8fafc;
            border-top: 2px solid #cbd5e1;
        }}
        .total-row td {{
            font-weight: 700;
            color: #0f172a;
            font-size: 13px;
            border-bottom: 2px solid #cbd5e1;
        }}
    </style>
    </head>
    <body>
    <div id="capture-container">
        <div class="header">
            <h2>Bảng theo dõi tồn backlog ODR TTS theo AM</h2>
            <p>NTB Region — Snapshot ngày {latest_date_display}</p>
        </div>
        <table>
            <thead>
                <tr>
                    <th style="width: 5%; text-align:center;">STT</th>
                    <th style="width: 25%;">AM phụ trách</th>
                    <th style="width: 11%; text-align:right;">Đã trễ ODR</th>
                    <th style="width: 13%; text-align:right;">Phải giao hôm nay</th>
                    <th style="width: 11%; text-align:right;">Còn 2 ngày</th>
                    <th style="width: 11%; text-align:right;">Còn 3 ngày</th>
                    <th style="width: 11%; text-align:right;">Còn >3 ngày</th>
                    <th style="width: 13%; text-align:right; background-color: #ea580c; color: white;">Tổng tồn backlog</th>
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
    
    # 5. Write HTML files
    t1_html_path = os.path.join(BASE_DIR, "t1_odr_rate.html")
    t2_html_path = os.path.join(BASE_DIR, "t2_odr_backlog.html")
    
    with open(t1_html_path, "w", encoding="utf-8") as f:
        f.write(t1_html)
    with open(t2_html_path, "w", encoding="utf-8") as f:
        f.write(t2_html)
        
    # 6. Screenshot using Playwright
    t1_image_path = os.path.join(BASE_DIR, "odr_rate_am.png")
    t2_image_path = os.path.join(BASE_DIR, "odr_backlog_am.png")
    
    print("📸 Rendering and capturing images with Playwright...")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        # Use a wide enough page viewport
        page = browser.new_page(viewport={"width": 1400, "height": 1200})
        
        # Capture Table 1
        page.goto(f"file:///{t1_html_path.replace('\\', '/')}")
        page.wait_for_timeout(1000)
        page.locator("#capture-container").screenshot(path=t1_image_path)
        print(f"✔️ Captured ODR Rates: {t1_image_path}")
        
        # Capture Table 2
        page.goto(f"file:///{t2_html_path.replace('\\', '/')}")
        page.wait_for_timeout(1000)
        page.locator("#capture-container").screenshot(path=t2_image_path)
        print(f"✔️ Captured ODR Backlog: {t2_image_path}")
        
        browser.close()
        
    # Clean up temp HTML files
    try:
        os.remove(t1_html_path)
        os.remove(t2_html_path)
    except:
        pass
        
    # Copy screenshots to artifact folder for user viewing if it exists
    artifact_dir = r"C:\Users\lap4all\.gemini\antigravity-ide\brain\cd8bf44c-b8bf-4b11-952a-d6a745f837cd"
    if os.path.exists(artifact_dir):
        import shutil
        try:
            shutil.copy2(t1_image_path, os.path.join(artifact_dir, "odr_rate_am.png"))
            shutil.copy2(t2_image_path, os.path.join(artifact_dir, "odr_backlog_am.png"))
            print("📋 Copied screenshots to brain artifact directory.")
        except Exception as e:
            print(f"⚠️ Warning copying screenshots to brain artifact: {e}")
            
    # 7. Post report and images to GTalk
    print("📡 Posting report and images to GTalk channel...")
    caption = f"""<b>BÁO CÁO ODR TTS & TỒN ODR TTS NTB NGÀY {latest_date_display}</b>

<b>1. Tỷ lệ ODR TTS toàn vùng NTB:</b> <b>{grand_odr_rate:.2f}%</b> {tot_diff_msg}
- GTC thành công đúng hạn: <b>{total_n1_ontime:,.0f} đơn</b>
- Tổng đơn GTC thành công: <b>{total_n1_can:,.0f} đơn</b>

<b>2. Tình hình tồn backlog ODR TTS:</b>
- Tổng tồn backlog hiện tại: <b>{sum_tot:,.0f} đơn</b>
  + Đã trễ ODR: <b>{sum_c1:,.0f} đơn</b> (<a href="https://docs.google.com/spreadsheets/d/{SHEET_KEY}/edit#gid=1077331166"><b>xem don den han giao</b></a>)
  + Phải giao hôm nay: <b>{sum_c2:,.0f} đơn</b>
  + Còn 2 ngày: <b>{sum_c3:,.0f} đơn</b>
  + Còn lại (3 ngày, >3 ngày): <b>{sum_c4 + sum_c5:,.0f} đơn</b>"""

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
        print("❌ Image upload to GTalk failed. img1_id:", img1_id, "img2_id:", img2_id)

    # Clean up local raw file and screenshots
    try:
        os.remove(t1_image_path)
        os.remove(t2_image_path)
    except:
        pass
        
    print("🎉 BÁO CÁO ODR & TỒN ODR TTS ĐÃ HOÀN THÀNH!")

if __name__ == "__main__":
    main()
