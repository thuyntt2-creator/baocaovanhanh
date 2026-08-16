# -*- coding: utf-8 -*-
"""
Script: analyze_am_aging_assignment_rate.py
Analyzes the assignment rates (assigned / active) of each AM across 4 aging brackets:
- 5 - 8 ngày (a, b, c)
- 8 - 10 ngày (d, e)
- 10 - 15 ngày (f, g, h, i, j)
- Trên 15 ngày (k)

Creates a styled dashboard tab 'Phân tích gán AM' in Google Sheets,
captures a screenshot with Playwright, and sends it to GTalk.
"""

import os
import sys
import json
import requests
import pandas as pd
import numpy as np
from datetime import datetime
import pytz
import gspread
from google.oauth2.service_account import Credentials
from google.oauth2.credentials import Credentials as UserCredentials
from playwright.sync_api import sync_playwright
import urllib3
from dotenv import load_dotenv

# Suppress SSL warnings
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Configure output encoding for Vietnamese characters
try:
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    if hasattr(sys.stderr, 'reconfigure'):
        sys.stderr.reconfigure(encoding='utf-8')
except Exception:
    pass

# ============ CONFIG & CONSTANTS ============
SHEET_KEY = '1WCzgao34cA_SttyB9ytHfE1qKTNl_3iFqDbEfw3lbyU'  # Target Spreadsheet
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')

# Load environment configuration if available
env_path = r"c:\Users\lap4all\Desktop\New folder\.env"
if os.path.exists(env_path):
    load_dotenv(dotenv_path=env_path, override=True)
else:
    load_dotenv()

GTALK_OA_TOKEN = os.environ.get("GTALK_OA_TOKEN") or "2067164759710552066:RfgqBJY4QtV18udu1wMEfhmoRNI4hgBv"
GTALK_CHANNEL_ID = os.environ.get("AGING_ASSIGNMENT_GTALK_CHANNEL_ID") or "2067164759710552066"

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

def merge_request(sheet_id, start_row, end_row, start_col, end_col):
    return {
        "mergeCells": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": start_row,
                "endRowIndex": end_row,
                "startColumnIndex": start_col,
                "endColumnIndex": end_col
            },
            "mergeType": "MERGE_ALL"
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

def col_width_request(sheet_id, start_col, end_col, width):
    return {
        "updateDimensionProperties": {
            "range": {
                "sheetId": sheet_id,
                "dimension": "COLUMNS",
                "startIndex": start_col,
                "endIndex": end_col
            },
            "properties": {
                "pixelSize": width
            },
            "fields": "pixelSize"
        }
    }

def border_request(sheet_id, start_row, end_row, start_col, end_col, color_hex="#BDC3C7"):
    color = make_color(color_hex)
    border_style = {
        "style": "SOLID",
        "color": color
    }
    return {
        "updateBorders": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": start_row,
                "endRowIndex": end_row,
                "startColumnIndex": start_col,
                "endColumnIndex": end_col
            },
            "top": border_style,
            "bottom": border_style,
            "left": border_style,
            "right": border_style,
            "innerHorizontal": border_style,
            "innerVertical": border_style
        }
    }

def get_group(col_k_val):
    val = str(col_k_val).strip().lower()
    if any(k in val for k in ['(a)', '(b)', '(c)']):
        return '5 - 8 ngày'
    elif any(k in val for k in ['(d)', '(e)']):
        return '>8 - 10 ngày'
    elif any(k in val for k in ['(f)', '(g)', '(h)', '(i)', '(j)']):
        return '>10 - 15 ngày'
    elif '(k)' in val:
        return 'Trên 15 ngày'
    return None

def main():
    # Only run in active hours unless --force is present
    tz = pytz.timezone("Asia/Ho_Chi_Minh")
    current_hour = datetime.now(tz).hour
    bypass_time = len(sys.argv) > 1 and sys.argv[1] == "--force"
    if not bypass_time and not (7 <= current_hour <= 18):
        print(f"💤 Ngoài khung giờ hoạt động (7h - 18h). Hiện tại là {datetime.now(tz).strftime('%H:%M:%S')}. Script sẽ dừng.")
        sys.exit(0)

    print(f"🚀 Bắt đầu phân tích tỷ lệ gán đơn aging lúc: {datetime.now(tz).strftime('%Y-%m-%d %H:%M:%S')}")

    # 1. Connect to Google Sheets
    gc_client = get_gspread_client(spreadsheet_id=SHEET_KEY)
    sh = gc_client.open_by_key(SHEET_KEY)

    print("📖 Đọc dữ liệu từ tab 'data LM'...")
    ws_lm = sh.worksheet("data LM")
    lm_data = ws_lm.get_all_values()

    if len(lm_data) < 2:
        print("❌ Lỗi: Sheet 'data LM' không đủ dữ liệu.")
        sys.exit(1)

    lm_header = lm_data[1]
    try:
        order_col_idx = lm_header.index("Mã đơn hàng")
        status_col_idx = lm_header.index("Trạng thái")
    except ValueError as e:
        print(f"❌ Lỗi: Không tìm thấy cột 'Mã đơn hàng' hoặc 'Trạng thái' trong 'data LM'. Chi tiết: {e}")
        sys.exit(1)

    # Map Mã đơn hàng -> Trạng thái
    lm_status = {}
    for row in lm_data[2:]:
        if len(row) > max(order_col_idx, status_col_idx):
            m_don = row[order_col_idx].strip()
            t_thai = row[status_col_idx].strip()
            if m_don:
                lm_status[m_don] = t_thai

    print("📖 Đọc dữ liệu từ tab 'Đơn giao aging trên 5 ngày'...")
    ws_aging = sh.worksheet("Đơn giao aging trên 5 ngày")
    aging_data = ws_aging.get_all_values()

    if len(aging_data) < 1:
        print("❌ Lỗi: Tab 'Đơn giao aging trên 5 ngày' trống.")
        sys.exit(1)

    aging_header = [h.strip().lower() for h in aging_data[0]]
    try:
        if "order_code" in aging_header:
            ag_order_idx = aging_header.index("order_code")
        else:
            ag_order_idx = aging_header.index("mã đơn")
        
        ag_am_idx = aging_header.index("am_name")
        
        ag_group_idx = -1
        for h_name in ["nhóm bl", "nhom bl", "nhóm"]:
            if h_name in aging_header:
                ag_group_idx = aging_header.index(h_name)
                break
    except ValueError as e:
        print(f"❌ Lỗi: Thiếu các cột bắt buộc. Chi tiết: {e}")
        sys.exit(1)

    # 2. Process and aggregate data
    stats = {} # am -> group -> {"active": x, "assigned": y}
    group_labels = ['5 - 8 ngày', '>8 - 10 ngày', '>10 - 15 ngày', 'Trên 15 ngày']

    for row in aging_data[1:]:
        if len(row) > max(ag_order_idx, ag_am_idx):
            m_don = row[ag_order_idx].strip()
            am = row[ag_am_idx].strip()
            if not m_don or not am:
                continue

            group_val = row[ag_group_idx].strip() if (ag_group_idx != -1 and len(row) > ag_group_idx) else ""
            group = get_group(group_val)
            if not group:
                continue

            # Check status from data LM
            status = lm_status.get(m_don, "đã giao/ chuyển trả thành công")
            # Filter out processed / completed orders
            is_processed = status in ['đã giao/ chuyển trả thành công', 'đã giao/chuyển trả thành công', '#n/a', 'n/a', 'thành công', '#N/A']
            if is_processed:
                continue

            # Initialize stats dict
            if am not in stats:
                stats[am] = {g: {"active": 0, "assigned": 0} for g in group_labels}

            # Check if assigned
            is_assigned = status in ["Đang có chuyến đi trong ngày", "Đã có chuyến đi trong ngày"]
            stats[am][group]["active"] += 1
            if is_assigned:
                stats[am][group]["assigned"] += 1

    # Filter active AMs and sort by total active count descending, then alphabetically by name
    active_ams = sorted(
        [am for am in stats.keys() if any(stats[am][g]["active"] > 0 for g in group_labels)],
        key=lambda am: (-sum(stats[am][g]["active"] for g in group_labels), am)
    )

    # Construct overall totals
    totals = {g: {"active": 0, "assigned": 0} for g in group_labels}
    for am in active_ams:
        for g in group_labels:
            totals[g]["active"] += stats[am][g]["active"]
            totals[g]["assigned"] += stats[am][g]["assigned"]

    # 3. Create or write to Phân tích gán AM Worksheet
    print("📊 Cập nhật dữ liệu vào Google Sheets...")
    ws_name = "Phân tích gán AM"
    try:
        ws = sh.worksheet(ws_name)
        ws.clear()
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=ws_name, rows="200", cols="10")

    # Layout construction
    now_str = datetime.now(tz).strftime('%Y-%m-%d %H:%M:%S')
    grid_values = []
    grid_values.append(["BẢNG PHÂN TÍCH TỶ LỆ GÁN ĐƠN AGING THEO AM", "", "", "", "", ""])
    grid_values.append([f"Cập nhật lúc: {now_str}", "", "", "", "", ""])
    grid_values.append([""]) # Empty divider row

    headers = ["AM"] + group_labels + ["Tổng cộng"]
    grid_values.append(headers)

    sheet_requests = []
    sheet_requests.append(merge_request(ws.id, 0, 1, 0, 6))
    sheet_requests.append(merge_request(ws.id, 1, 2, 0, 6))

    sheet_requests.append(cell_format_request(ws.id, 0, 1, 0, 6, {
        "textFormat": {"bold": True, "fontSize": 14, "fontFamily": "Arial", "foregroundColor": make_color("#1565C0")},
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "MIDDLE"
    }))
    sheet_requests.append(cell_format_request(ws.id, 1, 2, 0, 6, {
        "textFormat": {"fontSize": 10, "fontFamily": "Arial", "italic": True, "foregroundColor": make_color("#555555")},
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "MIDDLE"
    }))
    sheet_requests.append(cell_format_request(ws.id, 3, 4, 0, 6, {
        "backgroundColor": make_color("#1565C0"),
        "textFormat": {"bold": True, "fontSize": 10, "fontFamily": "Arial", "foregroundColor": make_color("#FFFFFF")},
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "MIDDLE"
    }))

    sheet_requests.append(row_height_request(ws.id, 0, 1, 35))
    sheet_requests.append(row_height_request(ws.id, 1, 2, 24))
    sheet_requests.append(row_height_request(ws.id, 3, 4, 30))

    # Helper function to compute and format cell text & requests
    def build_cell_text_and_format(assigned, active, r_idx, c_idx):
        if active == 0:
            text = "0/0 (—)"
            bg = "#ECEFF1"
            fg = "#78909C"
            bold = False
        else:
            rate = (assigned / active) * 100
            text = f"{assigned}/{active} ({rate:.0f}%)"
            bold = True
            if rate >= 90.0:
                bg = "#C8E6C9"
                fg = "#1B5E20"
            elif rate >= 70.0:
                bg = "#FFF9C4"
                fg = "#78350F"
            else:
                bg = "#FFCDD2"
                fg = "#B71C1C"
        
        req = cell_format_request(ws.id, r_idx, r_idx+1, c_idx, c_idx+1, {
            "backgroundColor": make_color(bg),
            "textFormat": {"foregroundColor": make_color(fg), "bold": bold, "fontFamily": "Arial", "fontSize": 9},
            "horizontalAlignment": "CENTER",
            "verticalAlignment": "MIDDLE"
        })
        return text, req

    data_start_row = 4
    for am_idx, am in enumerate(active_ams):
        row_idx = data_start_row + am_idx
        row_vals = [am]
        
        # AM name cell style
        sheet_requests.append(cell_format_request(ws.id, row_idx, row_idx+1, 0, 1, {
            "backgroundColor": make_color("#E3F2FD"),
            "textFormat": {"bold": True, "fontFamily": "Arial", "fontSize": 9},
            "horizontalAlignment": "LEFT",
            "verticalAlignment": "MIDDLE"
        }))
        sheet_requests.append(row_height_request(ws.id, row_idx, row_idx+1, 26))

        total_active = 0
        total_assigned = 0

        for g_idx, g in enumerate(group_labels):
            col_idx = 1 + g_idx
            assigned = stats[am][g]["assigned"]
            active = stats[am][g]["active"]
            
            total_active += active
            total_assigned += assigned

            text, req = build_cell_text_and_format(assigned, active, row_idx, col_idx)
            row_vals.append(text)
            sheet_requests.append(req)

        # Overall cell
        text, req = build_cell_text_and_format(total_assigned, total_active, row_idx, 5)
        row_vals.append(text)
        sheet_requests.append(req)

        grid_values.append(row_vals)

    # Add grand total row
    total_row_idx = data_start_row + len(active_ams)
    total_row_vals = ["TỔNG CỘNG"]

    sheet_requests.append(cell_format_request(ws.id, total_row_idx, total_row_idx+1, 0, 1, {
        "backgroundColor": make_color("#FFF176"),
        "textFormat": {"bold": True, "fontFamily": "Arial", "fontSize": 10},
        "horizontalAlignment": "LEFT",
        "verticalAlignment": "MIDDLE"
    }))
    sheet_requests.append(row_height_request(ws.id, total_row_idx, total_row_idx+1, 28))

    grand_active = 0
    grand_assigned = 0

    for g_idx, g in enumerate(group_labels):
        col_idx = 1 + g_idx
        assigned = totals[g]["assigned"]
        active = totals[g]["active"]
        
        grand_active += active
        grand_assigned += assigned

        text, req = build_cell_text_and_format(assigned, active, total_row_idx, col_idx)
        total_row_vals.append(text)
        sheet_requests.append(req)

    text, req = build_cell_text_and_format(grand_assigned, grand_active, total_row_idx, 5)
    total_row_vals.append(text)
    sheet_requests.append(req)

    grid_values.append(total_row_vals)

    # Apply borders and column widths
    sheet_requests.append(border_request(ws.id, 3, total_row_idx + 1, 0, 6))
    
    col_widths = {0: 160, 1: 120, 2: 120, 3: 120, 4: 120, 5: 120}
    for c_idx, w in col_widths.items():
        sheet_requests.append(col_width_request(ws.id, c_idx, c_idx+1, w))

    # Update sheet values
    ws.update(range_name=f"A1:F{len(grid_values)}", values=grid_values, value_input_option="USER_ENTERED")
    sh.batch_update({"requests": sheet_requests})
    print("✔️ Đã cập nhật xong sheet 'Phân tích gán AM'.")

    # 4. Generate visual HTML table & capture screenshot
    print("📸 Renderer: Vẽ bảng HTML và chụp hình...")
    
    def get_badge_html(assigned, active):
        if active == 0:
            return '<div class="rate-badge rate-empty">—</div><div class="count-sub">0 / 0</div>'
        rate = (assigned / active) * 100
        if rate >= 90.0:
            badge_class = "rate-green"
        elif rate >= 70.0:
            badge_class = "rate-yellow"
        else:
            badge_class = "rate-red"
        return f'<div class="rate-badge {badge_class}">{rate:.1f}%</div><div class="count-sub">{assigned} / {active}</div>'

    tbody_rows = ""
    for am in active_ams:
        am_total_active = sum(stats[am][g]["active"] for g in group_labels)
        am_total_assigned = sum(stats[am][g]["assigned"] for g in group_labels)
        
        tbody_rows += f"""
        <tr>
            <td class="am-name">{am}</td>
            <td>{get_badge_html(stats[am]['5 - 8 ngày']['assigned'], stats[am]['5 - 8 ngày']['active'])}</td>
            <td>{get_badge_html(stats[am]['>8 - 10 ngày']['assigned'], stats[am]['>8 - 10 ngày']['active'])}</td>
            <td>{get_badge_html(stats[am]['>10 - 15 ngày']['assigned'], stats[am]['>10 - 15 ngày']['active'])}</td>
            <td>{get_badge_html(stats[am]['Trên 15 ngày']['assigned'], stats[am]['Trên 15 ngày']['active'])}</td>
            <td style="background-color: #f8fafc; font-weight: 700;">
                {get_badge_html(am_total_assigned, am_total_active)}
            </td>
        </tr>
        """

    grand_total_row = f"""
    <tr class="total-row">
        <td>TỔNG CỘNG</td>
        <td>{get_badge_html(totals['5 - 8 ngày']['assigned'], totals['5 - 8 ngày']['active'])}</td>
        <td>{get_badge_html(totals['>8 - 10 ngày']['assigned'], totals['>8 - 10 ngày']['active'])}</td>
        <td>{get_badge_html(totals['>10 - 15 ngày']['assigned'], totals['>10 - 15 ngày']['active'])}</td>
        <td>{get_badge_html(totals['Trên 15 ngày']['assigned'], totals['Trên 15 ngày']['active'])}</td>
        <td style="background-color: #fef08a;">
            {get_badge_html(grand_assigned, grand_active)}
        </td>
    </tr>
    """

    html_content = f"""
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
            padding: 30px;
            display: flex;
            justify-content: center;
            align-items: center;
        }}
        #table-container {{
            background: #ffffff;
            padding: 32px;
            border-radius: 16px;
            box-shadow: 0 4px 20px rgba(15, 23, 42, 0.05);
            border: 1px solid #e2e8f0;
            width: 820px;
        }}
        .header {{
            margin-bottom: 24px;
            text-align: center;
            border-bottom: 2px solid #f1f5f9;
            padding-bottom: 20px;
        }}
        .header h2 {{
            margin: 0;
            font-family: 'Plus Jakarta Sans', sans-serif;
            font-size: 24px;
            color: #0f172a;
            font-weight: 800;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }}
        .header .subtitle {{
            font-size: 13px;
            color: #64748b;
            margin-top: 6px;
            font-weight: 500;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            text-align: center;
        }}
        th {{
            font-family: 'Plus Jakarta Sans', sans-serif;
            background-color: #f8fafc;
            color: #475569;
            font-weight: 700;
            font-size: 13px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            padding: 12px 6px;
            border-bottom: 2px solid #e2e8f0;
        }}
        td {{
            padding: 10px 6px;
            border-bottom: 1px solid #f1f5f9;
            vertical-align: middle;
        }}
        .am-name {{
            font-weight: 600;
            color: #0f172a;
            text-align: left;
            padding-left: 12px;
            background-color: #f8fafc;
            font-size: 14px;
        }}
        .rate-badge {{
            display: inline-block;
            padding: 4px 10px;
            border-radius: 9999px;
            font-size: 13px;
            font-weight: 700;
            min-width: 55px;
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
        .rate-empty {{
            background-color: #f1f5f9;
            color: #94a3b8;
        }}
        .count-sub {{
            font-size: 10px;
            color: #64748b;
            margin-top: 3px;
            font-weight: 500;
        }}
        .total-row td {{
            background-color: #fefcbf;
            font-weight: 800;
            color: #0f172a;
            border-top: 2px solid #cbd5e1;
            padding: 14px 6px;
            font-size: 14px;
        }}
    </style>
    </head>
    <body>
    <div id="table-container">
        <div class="header">
            <h2>Báo cáo Tỷ lệ gán đơn Aging theo AM</h2>
            <div class="subtitle">Mốc cập nhật: {now_str}</div>
        </div>
        <table>
            <thead>
                <tr>
                    <th style="text-align: left; padding-left: 12px; width: 140px;">AM</th>
                    <th style="width: 130px;">5 - 8 ngày</th>
                    <th style="width: 130px;">&gt;8 - 10 ngày</th>
                    <th style="width: 130px;">&gt;10 - 15 ngày</th>
                    <th style="width: 130px;">Trên 15 ngày</th>
                    <th style="width: 130px;">Tổng cộng</th>
                </tr>
            </thead>
            <tbody>
                {tbody_rows}
                {grand_total_row}
            </tbody>
        </table>
    </div>
    </body>
    </html>
    """

    temp_html_path = os.path.join(BASE_DIR, "temp_am_aging_analysis.html")
    with open(temp_html_path, "w", encoding="utf-8") as f:
        f.write(html_content)

    output_image_path = os.path.join(BASE_DIR, "table_am_aging_analysis.png")
    
    print("   Renderer: Chụp ảnh bằng Playwright...")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.goto(f"file:///{temp_html_path.replace('\\', '/')}")
        page.wait_for_timeout(1000)
        container = page.locator("#table-container")
        container.screenshot(path=output_image_path)
        browser.close()

    try:
        os.remove(temp_html_path)
    except:
        pass

    print(f"✔️ Ảnh đã lưu tại: {output_image_path}")

    # 5. Send message & upload photo to GTalk
    print("📡 Đang gửi ảnh báo cáo phân tích tỷ lệ gán sang GTalk group...")
    
    caption_text = (
        f"📊 <b>BÁO CÁO PHÂN TÍCH TỶ LỆ GÁN ĐƠN AGING THEO AM</b>\n"
        f"📅 Ngày: {datetime.now(tz).strftime('%d-%m-%Y')} | ⏱️ Mốc cập nhật: {datetime.now(tz).strftime('%H:%M')}\n"
        f"========================\n"
        f"🔗 Chi tiết Google Sheet: https://docs.google.com/spreadsheets/d/{SHEET_KEY}/edit#gid=0"
    )

    file_name = os.path.basename(output_image_path)
    file_size = os.path.getsize(output_image_path)
    
    with open(output_image_path, 'rb') as f:
        file_bytes = f.read()
        
    init_payload = {
        "ChannelId": GTALK_CHANNEL_ID,
        "FileName": file_name,
        "FileSize": str(file_size),
        "MimeType": "image/png",
        "Metadata": json.dumps({"width": 820, "height": 600}),
        "oaToken": GTALK_OA_TOKEN
    }
    
    try:
        resp_init = requests.post("https://mbff.ghn.vn/api/gtalk/initiate-upload", json=init_payload, timeout=20)
        if resp_init.status_code == 200:
            init_data = resp_init.json()
            if init_data.get("errorCode") == "success":
                presigned_url = init_data["data"]["PresignedURL"]
                upload_id = init_data["data"]["UploadId"]
                
                resp_put = requests.put(presigned_url, data=file_bytes, headers={"Content-Type": "image/png"}, timeout=40)
                if resp_put.status_code == 200:
                    resp_comp = requests.post("https://mbff.ghn.vn/api/gtalk/complete-upload", json={"oaToken": GTALK_OA_TOKEN, "UploadId": upload_id}, timeout=20)
                    if resp_comp.status_code == 200:
                        comp_data = resp_comp.json()
                        if comp_data.get("errorCode") == "success":
                            file_id = comp_data["data"]["Id"]
                            
                            send_payload = {
                                "channelId": GTALK_CHANNEL_ID,
                                "clientMsgId": str(int(datetime.now().timestamp() * 1000)),
                                "content": {
                                    "parseMode": "HTML",
                                    "attachment": {
                                        "caption": caption_text,
                                        "items": [{"image": {"fileId": file_id, "width": 820, "height": 600}}]
                                    }
                                },
                                "oaToken": GTALK_OA_TOKEN
                            }
                            resp_send = requests.post("https://mbff.ghn.vn/api/gtalk/send-message", json=send_payload, timeout=20)
                            if resp_send.status_code == 200:
                                send_res_data = resp_send.json()
                                if send_res_data.get("errorCode") == "success":
                                    print("✅ Đã gửi báo cáo gán AM lên GTalk group chat thành công!")
                                else:
                                    print(f"❌ Lỗi gửi tin nhắn GTalk (errorCode): {send_res_data}")
                            else:
                                print(f"❌ Lỗi gửi tin nhắn GTalk: {resp_send.status_code} - {resp_send.text}")
                        else:
                            print(f"❌ Lỗi hoàn tất upload GTalk (errorCode != success): {comp_data}")
                    else:
                        print(f"❌ Lỗi HTTP hoàn tất upload GTalk: {resp_comp.status_code} - {resp_comp.text}")
                else:
                    print(f"❌ Lỗi upload file lên presigned URL GTalk: {resp_put.status_code} - {resp_put.text}")
            else:
                print(f"❌ Lỗi khởi tạo upload GTalk (errorCode != success): {init_data}")
        else:
            print(f"❌ Lỗi HTTP khởi tạo upload GTalk: {resp_init.status_code} - {resp_init.text}")
    except Exception as e:
        print(f"❌ Lỗi khi gửi ảnh lên GTalk: {e}")
    
    # Also send to Telegram if configured
    tele_token = os.environ.get("TELEGRAM_BOT_TOKEN") or "8570130113:AAGXRiUaKBknVpgtm1_i9ZA47JRjAXmB21M"
    tele_chat = os.environ.get("TELEGRAM_CHAT_ID") or "-5058464865"
    print("📡 Đang gửi ảnh báo cáo phân tích tỷ lệ gán sang Telegram group...")
    tele_url = f"https://api.telegram.org/bot{tele_token}/sendPhoto"
    try:
        with open(output_image_path, 'rb') as f:
            resp_tele = requests.post(tele_url, data={
                "chat_id": tele_chat,
                "caption": caption_text,
                "parse_mode": "HTML"
            }, files={"photo": f}, timeout=15)
            if resp_tele.status_code == 200:
                print("✅ Đã gửi báo cáo sang Telegram group thành công!")
            else:
                print(f"⚠️ Telegram phản hồi lỗi: {resp_tele.status_code} - {resp_tele.text}")
    except Exception as e:
        print(f"⚠️ Lỗi khi gửi Telegram: {e}")

    print("🎉 HOÀN THÀNH TẤT CẢ!")

if __name__ == "__main__":
    main()
