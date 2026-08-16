# -*- coding: utf-8 -*-
"""
Script: push_am_tts_pickup.py
Reads TTS pickup orders from the 'raw' sheet of the OPR spreadsheet,
determines which are unassigned (no trip today), maps them to AMs,
creates individual worksheet tabs per AM containing their pending list,
renders a visual statistics table as a PNG image using Playwright,
and broadcasts both the image and custom links to GTalk.
"""

import os
import io
import sys
import json
import time
import requests
import gspread
import pandas as pd
import unicodedata
from google.oauth2.service_account import Credentials
from google.oauth2.credentials import Credentials as UserCredentials
from datetime import datetime
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

# Configure output encoding for Vietnamese characters
os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8', write_through=True)
    if hasattr(sys.stderr, 'reconfigure'):
        sys.stderr.reconfigure(encoding='utf-8', write_through=True)
except Exception:
    pass

# ============ CONFIG & CONSTANTS ============
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')

SPREADSHEET_ID = "1B-QCbEnPpILFFEWPYheGdmkgYV9gSf4lAyQMlhzwOCM"
GTALK_OA_TOKEN = os.environ.get("OPR_GTALK_OA_TOKEN") or "2067164759497973760:RfgqBJY4QtV18udu1wMEfhmoRNI4hgBv"
GTALK_CHANNEL_ID = os.environ.get("OPR_GTALK_CHANNEL_ID") or "2067283005274091520"

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

SERVICE_ACCOUNT_CANDIDATES = [
    os.path.join(BASE_DIR, 'credentials.json'),
    r'C:\Users\lap4all\Documents\Auto report\credentials.json',
    r'C:\Users\lap4all\Desktop\Backlog_Automation\credentials.json',
    r'C:\Users\lap4all\Downloads\credentials.json',
    'credentials.json',
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


def normalize_am_name(name):
    if not name or pd.isna(name):
        return "Không xác định"
    clean = " ".join(str(name).strip().split())
    if clean == "" or clean.lower() == "#n/a" or clean.lower() == "nan":
        return "Không xác định"
    
    prefix = ""
    name_part = clean
    dash_idx = clean.find("-")
    if dash_idx != -1:
        prefix_part = clean[:dash_idx]
        if prefix_part.isdigit():
            prefix = clean[:dash_idx + 1]
            name_part = clean[dash_idx + 1:]
            
    words = name_part.split(' ')
    normalized_words = []
    for w in words:
        if w.upper() == "AM":
            normalized_words.append("AM")
        elif len(w) > 0:
            normalized_words.append(w[0].upper() + w[1:].lower())
        else:
            normalized_words.append("")
    return prefix + " ".join(normalized_words)

# Helpers for sheet formatting
def make_color(hex_str):
    hex_str = hex_str.lstrip('#')
    return {
        "red": int(hex_str[0:2], 16) / 255.0,
        "green": int(hex_str[2:4], 16) / 255.0,
        "blue": int(hex_str[4:6], 16) / 255.0
    }

def cell_format_request(sheet_id, start_row, end_row, start_col, end_col, format_dict):
    return {
        "repeatCell": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": start_row,
                "endRowIndex": end_row,
                "startColumnIndex": start_col,
                "endColumnIndex": end_col
            },
            "cell": {
                "userEnteredFormat": format_dict
            },
            "fields": "userEnteredFormat(" + ",".join(format_dict.keys()) + ")"
        }
    }

def row_height_request(sheet_id, start_row, end_row, height):
    return {
        "updateDimensionProperties": {
            "range": {
                "sheetId": sheet_id,
                "dimension": "ROWS",
                "startIndex": start_row,
                "endIndex": end_row
            },
            "properties": {
                "pixelSize": height
            },
            "fields": "pixelSize"
        }
    }

def run_pipeline():
    print("🚀 BẮT ĐẦU QUY TRÌNH KIỂM TRA ĐƠN LẤY TTS LÚC:", datetime.now().strftime('%d/%m/%Y %H:%M:%S'))
    
    # 1. Authorize Google Sheet
    try:
        gc_client = get_gspread_client(spreadsheet_id=SPREADSHEET_ID)
        sh = gc_client.open_by_key(SPREADSHEET_ID)
        print(f"✔️ Đã kết nối thành công tới: '{sh.title}'")
    except Exception as e:
        print(f"❌ Lỗi kết nối Google Sheets: {e}")
        sys.exit(1)
        
    # 2. Read Worksheets
    try:
        ws_raw = sh.worksheet("raw")
        raw_values = ws_raw.get_all_records()
        df_raw = pd.DataFrame(raw_values)
        print(f"✔️ Đã đọc dữ liệu tab 'raw' ({len(df_raw)} dòng)")
    except Exception as e:
        print(f"❌ Lỗi đọc dữ liệu tab 'raw': {e}")
        sys.exit(1)
        
    try:
        ws_cocau = sh.worksheet("CoCauVung")
        cocau_data = ws_cocau.get_all_values()
        # Parse first 4 columns: warehouse_id, Bưu cục, Tỉnh, AM
        cocau_map = {}
        for row in cocau_data[1:]:
            if len(row) >= 4:
                w_id = str(row[0]).strip()
                bc_name = str(row[1]).strip()
                tinh = str(row[2]).strip()
                am = str(row[3]).strip()
                if w_id:
                    cocau_map[w_id] = {
                        "Bưu cục": bc_name,
                        "Tỉnh": tinh,
                        "AM": am
                    }
        print(f"✔️ Đã đọc dữ liệu tab 'CoCauVung' ({len(cocau_map)} mã kho)")
    except Exception as e:
        print(f"❌ Lỗi đọc dữ liệu tab 'CoCauVung': {e}")
        sys.exit(1)

    if df_raw.empty:
        print("⚠️ Tab 'raw' trống. Không có dữ liệu để xử lý.")
        sys.exit(0)
        
    # 3. Filter for: Loại đơn = "Lấy", Khách hàng = "TTS"
    df_raw['Loại đơn normalized'] = df_raw['Loại đơn'].astype(str).str.strip().str.lower()
    df_raw['Khách hàng normalized'] = df_raw['Khách hàng'].astype(str).str.strip().str.lower()
    
    df_tts_pickup = df_raw[
        (df_raw['Loại đơn normalized'] == 'lấy') & 
        (df_raw['Khách hàng normalized'] == 'tts')
    ]
    
    total_tts_pickup = len(df_tts_pickup)
    print(f"🎯 Tổng số đơn lấy TTS phát sinh: {total_tts_pickup}")
    
    if total_tts_pickup == 0:
        print("🎉 Không có đơn lấy TTS nào phát sinh hôm nay!")
        sys.exit(0)
        
    # 4. Map each order and build AM statistics
    # am_stats[am_name] = { "total": 0, "assigned": 0, "unassigned": 0 }
    am_stats = {}
    # am_unassigned_details[am_name] = [ [Order details] ]
    am_unassigned_details = {}
    # am_unassigned_bcs[am_name] = { buucuc_name: count }
    am_unassigned_bcs = {}
    
    for idx, row in df_tts_pickup.iterrows():
        poc_raw = str(row['Mã bưu cục']).strip()
        
        # Map post office info
        po_info = None
        if poc_raw in cocau_map:
            po_info = cocau_map[poc_raw]
        else:
            # Fallback prefix matching
            for cid, info in cocau_map.items():
                if cid.startswith(poc_raw) or poc_raw == cid:
                    po_info = info
                    break
                    
        if po_info:
            am_name = normalize_am_name(po_info["AM"])
            buucuc_name = po_info["Bưu cục"]
            tinh_name = po_info["Tỉnh"]
        else:
            am_name = "Không xác định"
            buucuc_name = f"Bưu cục {poc_raw}"
            tinh_name = "Không xác định"
            
        status = str(row['Trạng thái']).strip()
        
        is_unassigned = (status == "Chưa có chuyến đi trong ngày")
        is_assigned = status in ["Đang có chuyến đi trong ngày", "Đã có chuyến đi trong ngày"]
        
        if am_name not in am_stats:
            am_stats[am_name] = {"total": 0, "assigned": 0, "unassigned": 0}
            
        am_stats[am_name]["total"] += 1
        
        if is_assigned:
            am_stats[am_name]["assigned"] += 1
        elif is_unassigned:
            am_stats[am_name]["unassigned"] += 1
            
            # Store detail
            if am_name not in am_unassigned_details:
                am_unassigned_details[am_name] = []
            am_unassigned_details[am_name].append([
                row.get('Mã đơn hàng', ''),
                poc_raw,
                buucuc_name,
                tinh_name,
                status,
                row.get('Thời gian tồn đọng', ''),
                datetime.now().strftime('%d/%m/%Y %H:%M')
            ])
            
            # Store post office breakdown
            if am_name not in am_unassigned_bcs:
                am_unassigned_bcs[am_name] = {}
            if buucuc_name not in am_unassigned_bcs[am_name]:
                am_unassigned_bcs[am_name][buucuc_name] = 0
            am_unassigned_bcs[am_name][buucuc_name] += 1
            
    # Calculate totals
    grand_total = sum(s["total"] for s in am_stats.values())
    grand_assigned = sum(s["assigned"] for s in am_stats.values())
    grand_unassigned = sum(s["unassigned"] for s in am_stats.values())
    grand_rate = (grand_assigned / grand_total * 100.0) if grand_total > 0 else 0.0

    print(f"📊 Phân tích tổng thể: Tổng = {grand_total} | Đã gán = {grand_assigned} | Chưa gán = {grand_unassigned} | Tỷ lệ = {grand_rate:.2f}%")

    # 5. Clear old AM sheets and create new ones for active unassigned AMs
    print("🧹 Dọn dẹp các tab cũ...")
    try:
        all_sheets = sh.worksheets()
        for ws in all_sheets:
            if ws.title.startswith("[TTS Lấy]"):
                try:
                    sh.del_worksheet(ws)
                    print(f"  • Đã xóa tab cũ: {ws.title}")
                except Exception as e:
                    print(f"  • Không xóa được tab {ws.title}: {e}")
    except Exception as e:
        print(f"⚠️ Lỗi dọn dẹp các tab cũ: {e}")

    # Create new tabs and record sheet links
    am_sheet_links = {}
    unprocessed_headers = [
        'Mã đơn hàng', 'Mã bưu cục', 'Tên bưu cục', 
        'Tên tỉnh', 'Trạng thái hiện tại', 'Thời gian tồn đọng', 'Cập nhật lúc'
    ]

    for am_name, details in am_unassigned_details.items():
        if not details:
            continue
        tab_title = f"[TTS Lấy] {am_name}"
        # Keep title within Google Sheets limits (30 chars max is safe, tab_title might be longer but usually fit)
        if len(tab_title) > 30:
            tab_title = tab_title[:30]
            
        print(f"✍️ Đang tạo tab '{tab_title}' cho AM {am_name}...")
        try:
            # Sort detailed rows by Tên bưu cục
            details.sort(key=lambda x: x[2])
            
            ws_am = sh.add_worksheet(title=tab_title, rows=str(max(100, len(details) + 50)), cols="8")
            ws_am.update([unprocessed_headers] + details)
            
            # Format sheet header
            sh.batch_update({
                "requests": [
                    cell_format_request(ws_am.id, 0, 1, 0, len(unprocessed_headers), {
                        "backgroundColor": make_color("#F97316"), # Màu cam cảnh báo pickup
                        "textFormat": {"bold": True, "fontSize": 10, "fontFamily": "Arial", "foregroundColor": make_color("#FFFFFF")},
                        "horizontalAlignment": "CENTER",
                        "verticalAlignment": "MIDDLE"
                    }),
                    row_height_request(ws_am.id, 0, 1, 28)
                ]
            })
            am_sheet_links[am_name] = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit#gid={ws_am.id}"
            print(f"  • Đã tạo và định dạng thành công!")
        except Exception as e:
            print(f"  • Lỗi tạo tab cho {am_name}: {e}")
            am_sheet_links[am_name] = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit"

    # Also create/update a Master tab for all unassigned orders
    master_tab_name = "Đơn Lấy TTS Chưa Gán"
    master_sheet_link = ""
    all_unassigned_rows = []
    for am, details in am_unassigned_details.items():
        for d in details:
            # Insert AM name at index 3
            row_copy = list(d)
            row_copy.insert(3, am)
            all_unassigned_rows.append(row_copy)
            
    if all_unassigned_rows:
        try:
            print(f"✍️ Cập nhật tab Master '{master_tab_name}'...")
            all_main_worksheets = {ws.title: ws for ws in sh.worksheets()}
            if master_tab_name in all_main_worksheets:
                ws_master = all_main_worksheets[master_tab_name]
                ws_master.clear()
            else:
                ws_master = sh.add_worksheet(title=master_tab_name, rows=str(max(100, len(all_unassigned_rows) + 50)), cols="10")
                
            master_headers = [
                'Mã đơn hàng', 'Mã bưu cục', 'Tên bưu cục', 'AM', 
                'Tên tỉnh', 'Trạng thái hiện tại', 'Thời gian tồn đọng', 'Cập nhật lúc'
            ]
            # Sort master rows by AM, then by Bưu cục
            all_unassigned_rows.sort(key=lambda x: (x[3], x[2]))
            ws_master.update([master_headers] + all_unassigned_rows)
            
            sh.batch_update({
                "requests": [
                    cell_format_request(ws_master.id, 0, 1, 0, len(master_headers), {
                        "backgroundColor": make_color("#EA580C"), # Màu cam sậm
                        "textFormat": {"bold": True, "fontSize": 10, "fontFamily": "Arial", "foregroundColor": make_color("#FFFFFF")},
                        "horizontalAlignment": "CENTER",
                        "verticalAlignment": "MIDDLE"
                    }),
                    row_height_request(ws_master.id, 0, 1, 28)
                ]
            })
            master_sheet_link = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit#gid={ws_master.id}"
        except Exception as e:
            print(f"⚠️ Lỗi cập nhật tab master: {e}")

    # 6. Render Visual Table HTML
    now_str = datetime.now().strftime('%d/%m/%Y %H:%M')
    
    # Sort AM list for rendering: first by rate ascending (lowest first), then by unassigned descending
    sorted_am_stats = []
    for am, s in am_stats.items():
        rate = (s["assigned"] / s["total"] * 100.0) if s["total"] > 0 else 0.0
        sorted_am_stats.append((am, s["total"], s["assigned"], s["unassigned"], rate))
    sorted_am_stats.sort(key=lambda x: (x[4], -x[3]))

    # Generate HTML Rows
    table_rows_html = ""
    for idx, (am_name, tot, ass, unass, rate) in enumerate(sorted_am_stats, 1):
        # Determine Badge Style
        if rate == 100.0:
            badge_class = "badge-success"
        elif rate >= 80.0:
            badge_class = "badge-warning"
        else:
            badge_class = "badge-danger"
            
        unass_class = "pending-highlight" if unass > 0 else ""
        
        table_rows_html += f"""
        <tr>
            <td style="text-align: center; color: #64748b;">{idx}</td>
            <td class="am-name">{am_name}</td>
            <td class="number">{tot}</td>
            <td class="number" style="color: #10b981; font-weight: 600;">{ass}</td>
            <td class="number {unass_class}">{unass}</td>
            <td class="number">
                <span class="badge {badge_class}">{rate:.1f}%</span>
            </td>
        </tr>
        """

    html_content = f"""<!DOCTYPE html>
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
            box-shadow: 0 4px 24px rgba(15, 23, 42, 0.06);
            border: 1px solid #e2e8f0;
            width: 1000px;
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
            margin: 6px 0 0 0;
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
            padding: 14px 10px;
            border-bottom: 2px solid #e2e8f0;
        }}
        td {{
            padding: 12px 10px;
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
        .badge {{
            display: inline-block;
            padding: 4px 8px;
            border-radius: 9999px;
            font-size: 12px;
            font-weight: 600;
            text-align: center;
        }}
        .badge-success {{
            background-color: #dcfce7;
            color: #16a34a;
        }}
        .badge-warning {{
            background-color: #ffedd5;
            color: #ea580c;
        }}
        .badge-danger {{
            background-color: #fee2e2;
            color: #dc2626;
        }}
        .pending-highlight {{
            color: #dc2626;
            font-weight: 700;
            background-color: #fef2f2;
            border-radius: 4px;
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
            padding: 14px 10px;
        }}
    </style>
    </head>
    <body>
    <div id="capture-container">
        <div class="header">
            <h2>Bảng theo dõi tỷ lệ gán đơn lấy TTS theo AM</h2>
            <p>NTB Region — Cập nhật lúc {now_str}</p>
        </div>
        <table>
            <thead>
                <tr>
                    <th style="width: 5%; text-align:center;">STT</th>
                    <th style="width: 30%;">AM phụ trách</th>
                    <th style="width: 15%; text-align:right;">Tổng phát sinh</th>
                    <th style="width: 15%; text-align:right;">Đã gán chuyến</th>
                    <th style="width: 15%; text-align:right;">Chưa gán (Pending)</th>
                    <th style="width: 20%; text-align:right;">Tỷ lệ gán</th>
                </tr>
            </thead>
            <tbody>
                {table_rows_html}
                <tr class="total-row">
                    <td style="text-align: center;">-</td>
                    <td>TỔNG TOÀN VÙNG (GRAND TOTAL)</td>
                    <td class="number">{grand_total}</td>
                    <td class="number" style="color: #10b981;">{grand_assigned}</td>
                    <td class="number" style="color: { '#dc2626' if grand_unassigned > 0 else '#334155' };">{grand_unassigned}</td>
                    <td class="number">
                        <span class="badge { 'badge-success' if grand_rate == 100.0 else 'badge-warning' if grand_rate >= 80.0 else 'badge-danger' }">{grand_rate:.1f}%</span>
                    </td>
                </tr>
            </tbody>
        </table>
    </div>
    </body>
    </html>
    """

    # Write HTML file
    html_path = os.path.join(BASE_DIR, "temp_tts_pickup.html")
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html_content)

    # 7. Capture Image using Playwright
    image_path = os.path.join(BASE_DIR, "tts_pickup_am.png")
    print("📸 Đang chụp ảnh bảng số liệu bằng Playwright...")
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            page = browser.new_page(viewport={"width": 1200, "height": 1000})
            page.goto(f"file:///{html_path.replace('\\', '/')}")
            page.wait_for_timeout(1000)
            page.locator("#capture-container").screenshot(path=image_path)
            browser.close()
        print(f"✔️ Đã chụp thành công: {image_path}")
    except Exception as e:
        print(f"❌ Lỗi chụp ảnh Playwright: {e}")
        image_path = None

    # Clean up HTML file
    try:
        os.remove(html_path)
    except:
        pass

    # 8. Enrich GTalk Alert Caption
    caption = f"🚨 <b>ĐƠN LẤY TTS CHƯA GÁN CHUYẾN ĐI LẤY HÀNG</b>\n"
    caption += f"⏱️ <b>Mốc cập nhật:</b> {now_str}\n"
    if master_sheet_link:
        caption += f"🔗 <b>Tổng hợp toàn vùng:</b> <a href=\"{master_sheet_link}\"><b>Xem danh sách tất cả AM</b></a>\n"
        
    caption += f"\n📦 <b>TỔNG HỢP TÌNH TRẠNG ĐƠN LẤY TTS HÔM NAY:</b>\n"
    caption += f"  • <b>Tổng đơn lấy phát sinh:</b> <b>{grand_total}</b> đơn\n"
    caption += f"  • <b>Đã gán chuyến đi:</b> <b>{grand_assigned}</b> đơn ({grand_rate:.1f}%)\n"
    caption += f"  • <b>Chưa gán chuyến đi:</b> <b>{grand_unassigned}</b> đơn"

    print("--- NỘI DUNG CAPTION ---")
    print(caption)
    print("------------------------")

    # 9. Post message and image to GTalk
    img_id = None
    if image_path and os.path.exists(image_path):
        print("📡 Đang upload ảnh lên GTalk...")
        file_name = os.path.basename(image_path)
        file_size = os.path.getsize(image_path)
        with open(image_path, 'rb') as f:
            file_bytes = f.read()
            
        init_payload = {
            "ChannelId": GTALK_CHANNEL_ID,
            "FileName": file_name,
            "FileSize": str(file_size),
            "MimeType": "image/png",
            "Metadata": json.dumps({"width": 1000, "height": 800}),
            "oaToken": GTALK_OA_TOKEN
        }
        
        try:
            resp_init = requests.post("https://mbff.ghn.vn/api/gtalk/initiate-upload", json=init_payload, timeout=20, verify=False)
            if resp_init.status_code == 200:
                init_data = resp_init.json()
                if init_data.get("errorCode") == "success":
                    presigned_url = init_data["data"]["PresignedURL"]
                    upload_id = init_data["data"]["UploadId"]
                    
                    resp_put = requests.put(presigned_url, data=file_bytes, headers={"Content-Type": "image/png"}, timeout=60, verify=False)
                    if resp_put.status_code == 200:
                        resp_comp = requests.post("https://mbff.ghn.vn/api/gtalk/complete-upload", json={"oaToken": GTALK_OA_TOKEN, "UploadId": upload_id}, timeout=20, verify=False)
                        if resp_comp.status_code == 200:
                            comp_data = resp_comp.json()
                            if comp_data.get("errorCode") == "success":
                                img_id = comp_data["data"]["Id"]
                                print("✅ Đã upload ảnh thành công!")
        except Exception as e:
            print(f"❌ Lỗi khi upload ảnh lên GTalk: {e}")

    # Send message payload
    print("📡 Đang gửi thông điệp cảnh báo sang GTalk...")
    url_send = "https://mbff.ghn.vn/api/gtalk/send-message"
    client_msg_id = str(int(time.time() * 1000))
    
    if img_id:
        # Send with attachment
        payload = {
            "channelId": GTALK_CHANNEL_ID,
            "clientMsgId": client_msg_id,
            "content": {
                "parseMode": "HTML",
                "attachment": {
                    "caption": caption,
                    "items": [
                        {
                            "image": {
                                "fileId": img_id,
                                "width": 1000,
                                "height": 800
                            }
                        }
                    ]
                }
            },
            "oaToken": GTALK_OA_TOKEN
        }
    else:
        # Send as plain text if image upload failed
        payload = {
            "channelId": GTALK_CHANNEL_ID,
            "clientMsgId": client_msg_id,
            "content": {
                "parseMode": "HTML",
                "text": caption
            },
            "oaToken": GTALK_OA_TOKEN
        }

    try:
        res = requests.post(url_send, json=payload, headers={"Content-Type": "application/json"}, timeout=20, verify=False)
        if res.status_code == 200:
            res_data = res.json()
            if res_data.get("errorCode") == "success":
                print("✅ Đã gửi báo cáo sang GTalk group thành công!")
            else:
                print(f"❌ Gửi tin nhắn GTalk lỗi API: {res_data.get('error')}")
        else:
            print(f"❌ Gửi tin nhắn GTalk lỗi HTTP {res.status_code}: {res.text}")
    except Exception as e:
        print(f"❌ Lỗi kết nối gửi tin nhắn GTalk: {e}")

    # Clean up local image file
    if image_path:
        try:
            os.remove(image_path)
        except:
            pass
            
    print("🎉 HOÀN THÀNH QUY TRÌNH PUSH AM GÁN ĐƠN LẤY TTS!")

def main():
    try:
        run_pipeline()
    except Exception as e:
        import traceback
        now_str = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
        tb_str = traceback.format_exc()
        error_msg = f"⚠️ <b>[BOT LỖI - PUSH AM TTS PICKUP]</b>\n" \
                    f"⏱️ <b>Thời gian:</b> {now_str}\n" \
                    f"❌ <b>Lỗi ngắn:</b> <code>{str(e)}</code>\n\n" \
                    f"🛠️ <b>Traceback chi tiết:</b>\n<pre>{tb_str[:1500]}</pre>"
        
        print("❌ PHÁT HIỆN LỖI CRASH SCRIPT. ĐANG GỬI TIN BÁO LỖI NỘI BỘ...")
        print(tb_str)
        
        # Load env
        env_path = r"c:\Users\lap4all\Desktop\New folder\.env"
        if os.path.exists(env_path):
            load_dotenv(dotenv_path=env_path, override=True)
        else:
            load_dotenv()
            
        gtalk_admin = os.environ.get("OPR_GTALK_CHANNEL_ID_ADMIN") or os.environ.get("GTALK_CHANNEL_ID_ADMIN") or "2067164759710552066"
        
        if gtalk_admin:
            try:
                url_gtalk = "https://mbff.ghn.vn/api/gtalk/send-message"
                payload_gtalk = {
                    "channelId": gtalk_admin,
                    "clientMsgId": str(int(time.time() * 1000)),
                    "content": {
                        "parseMode": "HTML",
                        "text": error_msg
                    },
                    "oaToken": GTALK_OA_TOKEN
                }
                requests.post(url_gtalk, json=payload_gtalk, headers={"Content-Type": "application/json"}, timeout=10, verify=False)
                print("🟢 Đã gửi báo lỗi qua GTalk nội bộ.")
            except Exception as gtalk_err:
                print(f"❌ Không gửi được báo lỗi qua GTalk: {gtalk_err}")
                
        # Dọn dẹp file tạm
        try:
            if os.path.exists("temp_tts_pickup.html"):
                os.remove("temp_tts_pickup.html")
            if os.path.exists("tts_pickup_am.png"):
                os.remove("tts_pickup_am.png")
        except Exception:
            pass

        sys.exit(1)

if __name__ == "__main__":
    main()
