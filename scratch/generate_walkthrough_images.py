import os
import sys
import json
import pandas as pd
import gspread
import unicodedata
from datetime import datetime, timedelta
from google.oauth2.service_account import Credentials
from playwright.sync_api import sync_playwright

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
SHEET_KEY = '14r8n9L2cIG1Bmz8kSH79B24QzmnOApZhniGyOU40hr4'
ARTIFACTS_DIR = r"C:\Users\lap4all\.gemini\antigravity-ide\brain\3eed35b6-1a5a-4421-bcd4-8a8ddb149d05"

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
            print(f"Warning loading extracted_mappings.json: {e}")
    return {}

def resolve_po_info(po_name, cocau_map, std_mappings):
    po_norm = normalize_str(po_name)
    po_key = po_norm.lower()
    if po_key in cocau_map:
        return cocau_map[po_key]
    if po_key in std_mappings:
        return std_mappings[po_key]
    clean_po = po_key.replace("bưu cục", "").replace("bc", "").replace("bưu cục", "").replace(" ", "")
    for k, v in std_mappings.items():
        clean_k = k.replace("bưu cục", "").replace("bc", "").replace("bưu cục", "").replace(" ", "")
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
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except AttributeError:
        pass
        
    target_date = datetime(2026, 6, 22)
    date_str = target_date.strftime("%Y-%m-%d")
    print(f"Target date is N-1: {date_str}")
    
    credentials = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)

    gc_client = gspread.authorize(credentials)
    sh = gc_client.open_by_key(SHEET_KEY)
    
    # Read worksheets
    print("Reading Cơ cấu worksheet...")
    ws_cocau = sh.worksheet("Cơ cấu")
    cocau_rows = ws_cocau.get_all_values()
    df_cocau = pd.DataFrame(cocau_rows[1:], columns=cocau_rows[0])
    cocau_map = {normalize_str(row['Bưu cục']).lower(): (row['AM'], row['Tỉnh']) for idx, row in df_cocau.iterrows()}
    std_mappings = get_am_province_mappings()
    
    print("Reading data rớt LC worksheet...")
    ws_rot = sh.worksheet("data rớt LC")
    rot_rows = ws_rot.get_all_values()
    df_raw_ntb = pd.DataFrame(rot_rows[1:], columns=rot_rows[0])
    
    categories = ['Khác', 'Shopee', 'TTS']
    dfs_updated = {}
    
    for cat in categories:
        print(f"Reading Data {cat} worksheet...")
        ws_data = sh.worksheet(f"Data {cat}")
        data_rows = ws_data.get_all_values()
        df_data = pd.DataFrame(data_rows[1:], columns=data_rows[0])
        df_data['Vol cần LC'] = pd.to_numeric(df_data['Vol cần LC'], errors='coerce').fillna(0)
        df_data['%_rot_lc'] = pd.to_numeric(df_data['%_rot_lc'], errors='coerce').fillna(0)
        df_data['Vol rớt LC'] = df_data['Vol cần LC'] * df_data['%_rot_lc']
        dfs_updated[cat] = df_data
        
    n2_date_str = (target_date - timedelta(days=1)).strftime("%Y-%m-%d")
    
    # Check if we have N-2 data
    has_n2 = False
    for cat in categories:
        df_n2_cat = dfs_updated[cat][dfs_updated[cat]['Loại ngày'] == n2_date_str]
        if len(df_n2_cat) > 0 and df_n2_cat['Vol cần LC'].sum() > 0:
            has_n2 = True
            break
            
    # Calculate TTS N-1 metrics
    df_n1_tts = dfs_updated['TTS'][dfs_updated['TTS']['Loại ngày'] == date_str]
    n1_tts_can = df_n1_tts['Vol cần LC'].sum()
    n1_tts_rot = df_n1_tts['Vol rớt LC'].sum()
    n1_tts_rate = (n1_tts_rot / n1_tts_can * 100) if n1_tts_can > 0 else 0.0
    
    # Establish compatibility variables for Table 2 rendering
    df_n1 = df_n1_tts
    df_n2 = dfs_updated['TTS'][dfs_updated['TTS']['Loại ngày'] == n2_date_str]

    
    # 1. Capture Table 1: Top 10 POs
    df_raw_ntb['Grand_Total'] = 1
    po_counts = df_raw_ntb.groupby('tenbcxuat').agg(
        Khac=('loai_kh', lambda x: (x == 'Khac').sum()),
        Shopee=('loai_kh', lambda x: (x == 'Shopee').sum()),
        TTS=('loai_kh', lambda x: (x == 'TTS').sum()),
        Grand_Total=('Grand_Total', 'sum')
    ).reset_index()
    po_counts = po_counts.sort_values(by='Grand_Total', ascending=False).head(10).reset_index(drop=True)
    
    t1_rows_html = ""
    t1_sum_khac = t1_sum_shopee = t1_sum_tts = t1_sum_grand = 0
    for idx, row in po_counts.iterrows():
        po_name = row['tenbcxuat']
        am, _ = resolve_po_info(po_name, cocau_map, std_mappings)
        t1_rows_html += f"""
        <tr>
            <td class="number">{idx + 1}</td>
            <td class="po-name">{po_name}</td>
            <td class="am-name">{am}</td>
            <td class="number">{row['Khac']:,}</td>
            <td class="number">{row['Shopee']:,}</td>
            <td class="number">{row['TTS']:,}</td>
            <td class="number grand-total">{row['Grand_Total']:,}</td>
        </tr>
        """
        t1_sum_khac += row['Khac']
        t1_sum_shopee += row['Shopee']
        t1_sum_tts += row['TTS']
        t1_sum_grand += row['Grand_Total']
        
    t1_rows_html += f"""
    <tr class="total-row">
        <td colspan="3">TỔNG CỘNG TOP 10</td>
        <td class="number">{t1_sum_khac:,}</td>
        <td class="number">{t1_sum_shopee:,}</td>
        <td class="number">{t1_sum_tts:,}</td>
        <td class="number">{t1_sum_grand:,}</td>
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
            width: 1000px;
        }}
        .header {{
            margin-bottom: 20px;
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
            font-size: 14px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            padding: 12px 10px;
            border-bottom: 2px solid #e2e8f0;
        }}
        td {{
            padding: 10px 10px;
            font-size: 14px;
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
            padding: 14px 10px;
            font-size: 14px;
        }}
    </style>
    </head>
    <body>
    <div id="capture-container">
        <div class="header">
            <h2>Top 10 Bưu Cục Rớt Luân Chuyển Theo Loại Hàng</h2>
            <p>Khu vực Nam Trung Bộ | Ngày {target_date.strftime('%d/%m/%Y')}</p>
        </div>
        <table>
            <thead>
                <tr>
                    <th style="width: 5%; text-align:center;">STT</th>
                    <th style="width: 40%;">Chi tiết</th>
                    <th style="width: 20%;">AM</th>
                    <th style="width: 10%; text-align:center;">Khác</th>
                    <th style="width: 10%; text-align:center;">Shopee</th>
                    <th style="width: 10%; text-align:center;">TTS</th>
                    <th style="width: 10%; text-align:center;">Grand Total</th>
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
    
    # 2. Capture Table 2: AM rates
    df_n1_copy = df_n1.copy()
    df_n1_copy['AM_info'] = df_n1_copy['Chi tiết'].apply(lambda x: resolve_po_info(x, cocau_map, std_mappings))
    df_n1_copy['AM_name'] = df_n1_copy['AM_info'].apply(lambda x: x[0])
    df_n1_copy['Tỉnh_name'] = df_n1_copy['AM_info'].apply(lambda x: x[1])
    
    df_n2_copy = df_n2.copy()
    df_n2_copy['AM_info'] = df_n2_copy['Chi tiết'].apply(lambda x: resolve_po_info(x, cocau_map, std_mappings))
    df_n2_copy['AM_name'] = df_n2_copy['AM_info'].apply(lambda x: x[0])
    
    am_n2_rates = df_n2_copy.groupby('AM_name').agg(
        n2_can=('Vol cần LC', 'sum'),
        n2_rot=('Vol rớt LC', 'sum')
    ).reset_index()
    am_n2_rates['n2_rate'] = am_n2_rates['n2_rot'] / am_n2_rates['n2_can']
    am_n2_map = dict(zip(am_n2_rates['AM_name'], am_n2_rates['n2_rate']))
    
    am_n1_rates = df_n1_copy.groupby(['AM_name', 'Tỉnh_name']).agg(
        n1_can=('Vol cần LC', 'sum'),
        n1_rot=('Vol rớt LC', 'sum')
    ).reset_index()
    am_n1_rates['n1_rate'] = am_n1_rates['n1_rot'] / am_n1_rates['n1_can']
    am_n1_rates = am_n1_rates.sort_values(by='n1_rate', ascending=False).reset_index(drop=True)
    
    t2_rows_html = ""
    for idx, row in am_n1_rates.iterrows():
        am_name = row['AM_name']
        prov_name = row['Tỉnh_name']
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
            
        if has_n2:
            prev_rate = am_n2_map.get(am_name, 0.0) * 100
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
            <td class="province-name">{prov_name}</td>
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
    if has_n2:
        tot_diff = n1_tts_rate - n2_tts_rate
        if tot_diff > 0:
            tot_diff_html = f'<span class="compare-up">▲ +{tot_diff:.2f}%</span>'
        elif tot_diff < 0:
            tot_diff_html = f'<span class="compare-down">▼ {tot_diff:.2f}%</span>'
        else:
            tot_diff_html = '<span class="compare-same">—</span>'
    else:
        tot_diff_html = '<span class="compare-same">N/A</span>'
        
    t2_rows_html += f"""
    <tr class="total-row">
        <td colspan="3">TỔNG CỘNG</td>
        <td class="number">{n1_tts_can:,.0f}</td>
        <td class="number">{n1_tts_rot:,.0f}</td>
        <td style="text-align:center;">
            <span class="rate-badge {t2_grand_badge}"><span class="dot {t2_grand_dot}"></span>{n1_tts_rate:.2f}%</span>
        </td>
        <td style="text-align:center;">{tot_diff_html}</td>
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
            width: 900px;
        }}
        .header {{
            margin-bottom: 20px;
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
            font-size: 14px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            padding: 12px 10px;
            border-bottom: 2px solid #e2e8f0;
        }}
        td {{
            padding: 10px 10px;
            font-size: 14px;
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
        .province-name {{
            color: #475569;
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
            font-size: 13px;
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
            color: #b91c1c;
            font-weight: bold;
        }}
        .compare-down {{
            color: #15803d;
            font-weight: bold;
        }}
        .compare-same {{
            color: #64748b;
        }}
        .total-row td {{
            background-color: #fef2f2 !important;
            font-weight: 800;
            color: #b91c1c;
            border-top: 2px solid #ef4444;
            border-bottom: none;
            padding: 14px 10px;
            font-size: 14px;
        }}
    </style>
    </head>
    <body>
    <div id="capture-container">
        <div class="header">
            <h2>Báo Cáo Tỷ Lệ AM Rớt Luân Chuyển (TTS)</h2>
            <p>Khu vực Nam Trung Bộ | Ngày {target_date.strftime('%d/%m/%Y')}</p>
        </div>
        <table>
            <thead>
                <tr>
                    <th style="width: 5%; text-align:center;">STT</th>
                    <th style="width: 25%;">AM</th>
                    <th style="width: 20%;">Tỉnh</th>
                    <th style="width: 13%; text-align:center;">Vol cần LC</th>
                    <th style="width: 13%; text-align:center;">Vol rớt LC</th>
                    <th style="width: 12%; text-align:center;">Tỷ lệ rớt</th>
                    <th style="width: 12%; text-align:center;">So với hôm qua</th>
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
    
    t1_path = os.path.join(BASE_DIR, "scratch", "t1_walkthrough.html")
    t2_path = os.path.join(BASE_DIR, "scratch", "t2_walkthrough.html")
    with open(t1_path, "w", encoding="utf-8") as f:
        f.write(t1_html)
    with open(t2_path, "w", encoding="utf-8") as f:
        f.write(t2_html)
        
    t1_img = os.path.join(ARTIFACTS_DIR, "top10_po_rot_lc.png")
    t2_img = os.path.join(ARTIFACTS_DIR, "am_rates_rot_lc.png")
    
    print("Capturing screenshots with Playwright...")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        
        # Screenshot Table 1
        page.goto(f"file:///{t1_path.replace('\\', '/')}")
        page.wait_for_timeout(1000)
        page.locator("#capture-container").screenshot(path=t1_img)
        print(f"Captured Table 1: {t1_img}")
        
        # Screenshot Table 2
        page.goto(f"file:///{t2_path.replace('\\', '/')}")
        page.wait_for_timeout(1000)
        page.locator("#capture-container").screenshot(path=t2_img)
        print(f"Captured Table 2: {t2_img}")
        
        browser.close()
        
    try:
        os.remove(t1_path)
        os.remove(t2_path)
    except:
        pass
        
    print("Walkthrough images generated successfully!")

if __name__ == '__main__':
    main()
