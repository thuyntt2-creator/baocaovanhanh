# -*- coding: utf-8 -*-
"""
Script: analyze_unattempted_aging_by_bracket.py
Analyzes active aging orders that have 0 delivery attempts (num_deliver == 0)
across 4 aging brackets:
- 5 - 8 ngày (a, b, c)
- >8 - 10 ngày (d, e)
- >10 - 15 ngày (f, g, h, i, j)
- Trên 15 ngày (k)

Creates a styled dashboard tab 'Phân tích chưa giao' and a detail tab 'Chi tiết AM xử lý' in Google Sheets,
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
from playwright.sync_api import sync_playwright
import urllib3
import unicodedata

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

GTALK_OA_TOKEN = "2067164759710552066:RfgqBJY4QtV18udu1wMEfhmoRNI4hgBv"
GTALK_CHANNEL_ID = "2067164759710552066"

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

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
    # Check active hours
    tz = pytz.timezone("Asia/Ho_Chi_Minh")
    current_hour = datetime.now(tz).hour
    bypass_time = len(sys.argv) > 1 and sys.argv[1] == "--force"
    if not bypass_time and not (7 <= current_hour <= 18):
        print(f"💤 Ngoài khung giờ hoạt động (7h - 18h). Hiện tại là {datetime.now(tz).strftime('%H:%M:%S')}. Script sẽ dừng.")
        sys.exit(0)

    print(f"🚀 Bắt đầu phân tích đơn aging chưa đi giao lúc: {datetime.now(tz).strftime('%Y-%m-%d %H:%M:%S')}")

    # 1. Connect to Google Sheets
    if not os.path.exists(JSON_FILE):
        print(f"❌ Không tìm thấy file credentials.json tại {JSON_FILE}")
        sys.exit(1)

    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    sh = gc_client.open_by_key(SHEET_KEY)

    print("📖 Đọc dữ liệu từ tab 'data LM'...")
    ws_lm = sh.worksheet("data LM")
    lm_data = ws_lm.get_all_values()

    print("📖 Đọc dữ liệu từ tab 'Đơn giao aging trên 5 ngày'...")
    ws_aging = sh.worksheet("Đơn giao aging trên 5 ngày")
    aging_data = ws_aging.get_all_values()

    print("📖 Đọc dữ liệu từ tab 'Cơ cấu'...")
    ws_cocau = sh.worksheet("Cơ cấu")
    cocau_data = ws_cocau.get_all_values()

    if len(lm_data) < 2 or len(aging_data) < 1:
        print("❌ Lỗi: Dữ liệu sheet trống hoặc không đủ dòng.")
        sys.exit(1)

    # Map Cơ cấu
    cocau_map = {}
    for r in cocau_data[1:]:
        if len(r) >= 4:
            id_bc = r[0].strip()
            bc_name = r[1].strip()
            am_name = unicodedata.normalize('NFC', r[3].strip())
            if id_bc:
                cocau_map[id_bc] = am_name
            if bc_name:
                cocau_map[bc_name] = am_name

    # Extract backlog status from data LM
    lm_header = lm_data[1]
    try:
        lm_order_col = lm_header.index("Mã đơn hàng")
        lm_status_col = lm_header.index("Trạng thái")
    except ValueError as e:
        print(f"❌ Lỗi: Không tìm thấy cột Mã đơn hàng / Trạng thái trong tab data LM. Chi tiết: {e}")
        sys.exit(1)

    lm_status = {}
    for row in lm_data[2:]:
        if len(row) > max(lm_order_col, lm_status_col):
            m_don = row[lm_order_col].strip()
            t_thai = row[lm_status_col].strip()
            if m_don:
                lm_status[m_don] = t_thai

    # Process aging data headers
    aging_header = aging_data[0]
    try:
        ag_order_col = aging_header.index("order_code") if "order_code" in aging_header else aging_header.index("mã đơn")
        ag_bc_col = aging_header.index("bc")
        ag_id_bc_col = aging_header.index("id_bc")
        ag_group_col = aging_header.index("Nhóm BL")
        ag_am_col = aging_header.index("am_name")
        ag_num_col = aging_header.index("num_deliver") if "num_deliver" in aging_header else -1
    except ValueError as e:
        print(f"❌ Lỗi: Không tìm thấy cột bắt buộc trong tab Đơn giao aging trên 5 ngày. Chi tiết: {e}")
        sys.exit(1)

    SNAPSHOT_FILE = os.path.join(BASE_DIR, 'unattempted_snapshot.json')
    state = {"last_updated_date": "", "history": []}
    if os.path.exists(SNAPSHOT_FILE):
        try:
            with open(SNAPSHOT_FILE, "r", encoding="utf-8") as f:
                state = json.load(f)
        except Exception as e:
            print(f"⚠️ Cảnh báo đọc snapshot cũ: {e}")
            
    today_str = datetime.now(tz).strftime('%Y-%m-%d')
    current_time = datetime.now(tz).strftime('%H:%M')
    current_snap_orders = {}

    # 2. Process and aggregate unattempted orders
    success_keywords = ['đã giao/ chuyển trả thành công', 'đã giao/chuyển trả thành công', 'n/a', 'thành công']
    group_labels = ['5 - 8 ngày', '>8 - 10 ngày', '>10 - 15 ngày', 'Trên 15 ngày']
    stats = {} # am_name -> group -> {"unattempted": x, "active": y}

    for row in aging_data[1:]:
        if len(row) > max(ag_order_col, ag_bc_col, ag_id_bc_col, ag_group_col, ag_am_col):
            order_code = row[ag_order_col].strip()
            if not order_code:
                continue

            # Check if active
            status = lm_status.get(order_code, '#N/A')
            is_processed = status in ['#N/A', 'n/a'] or any(sk in status.lower() for sk in success_keywords)
            if is_processed:
                continue

            # Get group
            group_val = row[ag_group_col].strip()
            group = get_group(group_val)
            if not group:
                continue

            # Parse num_deliver to check if attempts == 0
            num_deliver_val = row[ag_num_col].strip() if (ag_num_col != -1 and len(row) > ag_num_col) else '0'
            try:
                num_attempts = int(float(num_deliver_val))
            except ValueError:
                num_attempts = 0

            # Match AM name with fallback
            raw_am = row[ag_am_col].strip()
            if not raw_am or raw_am == '#N/A' or raw_am == '':
                bc_name = row[ag_bc_col].strip()
                id_bc = row[ag_id_bc_col].strip()
                am_name = cocau_map.get(id_bc, cocau_map.get(bc_name, "Không xác định"))
            else:
                am_name = raw_am
            am_name = unicodedata.normalize('NFC', am_name)

            if am_name not in stats:
                stats[am_name] = {g: {"unattempted": 0, "active": 0} for g in group_labels}

            # Increment counts
            stats[am_name][group]["active"] += 1
            if num_attempts == 0:
                stats[am_name][group]["unattempted"] += 1

            current_snap_orders[order_code] = {
                "num_deliver": num_attempts,
                "am": am_name,
                "group": group
            }

    current_snap = {
        "time": current_time,
        "orders": current_snap_orders
    }

    # Compare current run with historical snapshot to detect processed orders
    compare_lines = []
    processed_records = []
    now_str = datetime.now(tz).strftime('%Y-%m-%d %H:%M:%S')
    
    if state.get("history"):
        if state.get("last_updated_date") != today_str:
            ref_snap = state["history"][-1]
            ref_label = "hôm qua"
        else:
            ref_snap = state["history"][0]
            ref_label = f"mốc sáng ({ref_snap['time']})"
            
        ref_orders = ref_snap.get("orders", {})
        processed_by_am = {}
        
        for o_code, info in ref_orders.items():
            if info.get("num_deliver", 0) == 0:
                prev_am = info.get("am", "Không xác định")
                is_still_active = o_code in current_snap_orders
                is_processed = False
                new_num = 0
                current_status = ""
                
                if not is_still_active:
                    is_processed = True
                    current_status = lm_status.get(o_code, "đã giao/ chuyển trả thành công")
                    new_num = 1
                else:
                    current_num = current_snap_orders[o_code]["num_deliver"]
                    if current_num >= 1:
                        is_processed = True
                        current_status = lm_status.get(o_code, "Chưa có chuyến đi trong ngày")
                        new_num = current_num
                        
                if is_processed:
                    processed_by_am[prev_am] = processed_by_am.get(prev_am, 0) + 1
                    order_group = info.get("group", "Không xác định")
                    
                    processed_records.append({
                        "code": o_code,
                        "am": prev_am,
                        "group": order_group,
                        "status": current_status,
                        "old_deliver": 0,
                        "new_deliver": new_num,
                        "ref_label": ref_label,
                        "recorded_time": now_str
                    })
                    
        if processed_by_am:
            compare_lines.append(f"📈 <b>Ghi nhận AM đã xử lý đơn chưa giao (so với {ref_label}):</b>")
            for am in sorted(processed_by_am.keys(), key=lambda x: processed_by_am[x], reverse=True):
                compare_lines.append(f"• AM <b>{am}</b>: đã xử lý <b>{processed_by_am[am]}</b> đơn")
            compare_lines.append("<i>(Chi tiết danh sách mã đơn xem tại tab 'Chi tiết AM xử lý')</i>")
        else:
            compare_lines.append(f"📈 <b>Ghi nhận AM đã xử lý đơn chưa giao (so với {ref_label}):</b>")
            compare_lines.append("• Không ghi nhận đơn nào từ chưa giao chuyển sang đã giao/hoàn thành.")
    else:
        compare_lines.append("📈 <b>Ghi nhận AM đã xử lý đơn chưa giao (so với hôm qua):</b>")
        compare_lines.append("• Đang khởi tạo dữ liệu đối chiếu lần đầu (chưa có mốc so sánh).")
        
    compare_text = "\n".join(compare_lines)

    # 2.7 Write details to 'Chi tiết AM xử lý' tab
    print("📊 Cập nhật chi tiết đơn đã xử lý vào Google Sheets...")
    detail_ws_name = "Chi tiết AM xử lý"
    try:
        detail_ws = sh.worksheet(detail_ws_name)
        detail_ws.clear()
    except gspread.exceptions.WorksheetNotFound:
        detail_ws = sh.add_worksheet(title=detail_ws_name, rows="1000", cols="10")
        
    detail_grid = [
        ["Mã đơn hàng", "AM phụ trách", "Nhóm tuổi đơn", "Trạng thái hiện tại", "Số lần giao cũ", "Số lần giao mới", "Mốc đối chiếu", "Thời điểm ghi nhận"]
    ]
    for rec in processed_records:
        detail_grid.append([
            rec["code"],
            rec["am"],
            rec["group"],
            rec["status"],
            rec["old_deliver"],
            rec["new_deliver"],
            rec["ref_label"],
            rec["recorded_time"]
        ])
        
    detail_ws.update(range_name=f"A1:H{len(detail_grid)}", values=detail_grid, value_input_option="USER_ENTERED")
    
    detail_requests = [
        cell_format_request(detail_ws.id, 0, 1, 0, 8, {
            "backgroundColor": make_color("#1565C0"),
            "textFormat": {"bold": True, "fontSize": 10, "fontFamily": "Arial", "foregroundColor": make_color("#FFFFFF")},
            "horizontalAlignment": "CENTER",
            "verticalAlignment": "MIDDLE"
        }),
        row_height_request(detail_ws.id, 0, 1, 28),
        border_request(detail_ws.id, 0, len(detail_grid), 0, 8)
    ]
    detail_col_widths = {0: 130, 1: 160, 2: 130, 3: 180, 4: 100, 5: 100, 6: 150, 7: 150}
    for c_idx, w in detail_col_widths.items():
        detail_requests.append(col_width_request(detail_ws.id, c_idx, c_idx+1, w))
        
    sh.batch_update({"requests": detail_requests})
    print("✔️ Đã cập nhật xong sheet 'Chi tiết AM xử lý'.")

    # Update state history and save
    if state.get("last_updated_date") != today_str:
        state["history"] = [current_snap]
        state["last_updated_date"] = today_str
    else:
        if len(state["history"]) == 1:
            state["history"].append(current_snap)
        elif len(state["history"]) > 1:
            state["history"][1] = current_snap
            
    try:
        with open(SNAPSHOT_FILE, "w", encoding="utf-8") as f:
            json.dump(state, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"⚠️ Cảnh báo không thể ghi snapshot: {e}")

    # Filter active AMs and sort by total unattempted count descending, then alphabetically by name
    active_ams = sorted(
        [am for am in stats.keys() if any(stats[am][g]["active"] > 0 for g in group_labels)],
        key=lambda am: (-sum(stats[am][g]["unattempted"] for g in group_labels), am)
    )

    # Construct overall totals
    totals = {g: {"unattempted": 0, "active": 0} for g in group_labels}
    for am in active_ams:
        for g in group_labels:
            totals[g]["active"] += stats[am][g]["active"]
            totals[g]["unattempted"] += stats[am][g]["unattempted"]

    # 3. Create or write to Phân tích chưa giao tab
    print("📊 Cập nhật dữ liệu vào Google Sheets...")
    ws_name = "Phân tích chưa giao"
    try:
        ws = sh.worksheet(ws_name)
        ws.clear()
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=ws_name, rows="200", cols="10")

    grid_values = []
    grid_values.append(["BẢNG PHÂN TÍCH ĐƠN AGING CHƯA ĐI GIAO LẦN NÀO THEO AM", "", "", "", "", ""])
    grid_values.append([f"Cập nhật lúc: {now_str}", "", "", "", "", ""])
    grid_values.append([""]) # Spacer row

    headers = ["AM"] + group_labels + ["Tổng cộng"]
    grid_values.append(headers)

    sheet_requests = []
    sheet_requests.append(merge_request(ws.id, 0, 1, 0, 6))
    sheet_requests.append(merge_request(ws.id, 1, 2, 0, 6))

    sheet_requests.append(cell_format_request(ws.id, 0, 1, 0, 6, {
        "textFormat": {"bold": True, "fontSize": 14, "fontFamily": "Arial", "foregroundColor": make_color("#B71C1C")},
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "MIDDLE"
    }))
    sheet_requests.append(cell_format_request(ws.id, 1, 2, 0, 6, {
        "textFormat": {"fontSize": 10, "fontFamily": "Arial", "italic": True, "foregroundColor": make_color("#555555")},
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "MIDDLE"
    }))
    sheet_requests.append(cell_format_request(ws.id, 3, 4, 0, 6, {
        "backgroundColor": make_color("#B71C1C"),
        "textFormat": {"bold": True, "fontSize": 10, "fontFamily": "Arial", "foregroundColor": make_color("#FFFFFF")},
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "MIDDLE"
    }))

    sheet_requests.append(row_height_request(ws.id, 0, 1, 35))
    sheet_requests.append(row_height_request(ws.id, 1, 2, 24))
    sheet_requests.append(row_height_request(ws.id, 3, 4, 30))

    def build_cell_text_and_format(unattempted, active, r_idx, c_idx):
        if active == 0:
            text = "0/0"
            bg = "#ECEFF1"
            fg = "#78909C"
            bold = False
        else:
            text = f"{unattempted}/{active}"
            bold = True
            if unattempted == 0:
                bg = "#C8E6C9"
                fg = "#1B5E20"
            elif unattempted <= 2:
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
        
        # AM name style
        sheet_requests.append(cell_format_request(ws.id, row_idx, row_idx+1, 0, 1, {
            "backgroundColor": make_color("#FADBD8"),
            "textFormat": {"bold": True, "fontFamily": "Arial", "fontSize": 9},
            "horizontalAlignment": "LEFT",
            "verticalAlignment": "MIDDLE"
        }))
        sheet_requests.append(row_height_request(ws.id, row_idx, row_idx+1, 26))

        total_active = 0
        total_unattempted = 0

        for g_idx, g in enumerate(group_labels):
            col_idx = 1 + g_idx
            unattempted = stats[am][g]["unattempted"]
            active = stats[am][g]["active"]
            
            total_active += active
            total_unattempted += unattempted

            text, req = build_cell_text_and_format(unattempted, active, row_idx, col_idx)
            row_vals.append(text)
            sheet_requests.append(req)

        # Overall totals for the AM
        text, req = build_cell_text_and_format(total_unattempted, total_active, row_idx, 5)
        row_vals.append(text)
        sheet_requests.append(req)

        grid_values.append(row_vals)

    # Grand total row
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
    grand_unattempted = 0

    for g_idx, g in enumerate(group_labels):
        col_idx = 1 + g_idx
        unattempted = totals[g]["unattempted"]
        active = totals[g]["active"]
        
        grand_active += active
        grand_unattempted += unattempted

        text, req = build_cell_text_and_format(unattempted, active, total_row_idx, col_idx)
        total_row_vals.append(text)
        sheet_requests.append(req)

    text, req = build_cell_text_and_format(grand_unattempted, grand_active, total_row_idx, 5)
    total_row_vals.append(text)
    sheet_requests.append(req)

    grid_values.append(total_row_vals)

    # Grid borders and sizes
    sheet_requests.append(border_request(ws.id, 3, total_row_idx + 1, 0, 6))
    
    col_widths = {0: 160, 1: 120, 2: 120, 3: 120, 4: 120, 5: 120}
    for c_idx, w in col_widths.items():
        sheet_requests.append(col_width_request(ws.id, c_idx, c_idx+1, w))

    # Update sheet content
    ws.update(range_name=f"A1:F{len(grid_values)}", values=grid_values, value_input_option="USER_ENTERED")
    sh.batch_update({"requests": sheet_requests})
    print("✔️ Đã cập nhật xong sheet 'Phân tích chưa giao'.")

    # 4. Generate visual HTML table & capture screenshot
    print("📸 Renderer: Vẽ bảng HTML chưa giao và chụp hình...")
    
    def get_badge_html(unattempted, active):
        if active == 0:
            return '<div class="rate-badge rate-empty">—</div><div class="count-sub">0 / 0</div>'
        if unattempted == 0:
            badge_class = "rate-green"
        elif unattempted <= 2:
            badge_class = "rate-yellow"
        else:
            badge_class = "rate-red"
        return f'<div class="rate-badge {badge_class}">{unattempted}</div><div class="count-sub">tồn {active} đơn</div>'

    tbody_rows = ""
    for am in active_ams:
        am_total_active = sum(stats[am][g]["active"] for g in group_labels)
        am_total_unattempted = sum(stats[am][g]["unattempted"] for g in group_labels)
        
        tbody_rows += f"""
        <tr>
            <td class="am-name">{am}</td>
            <td>{get_badge_html(stats[am]['5 - 8 ngày']['unattempted'], stats[am]['5 - 8 ngày']['active'])}</td>
            <td>{get_badge_html(stats[am]['>8 - 10 ngày']['unattempted'], stats[am]['>8 - 10 ngày']['active'])}</td>
            <td>{get_badge_html(stats[am]['>10 - 15 ngày']['unattempted'], stats[am]['>10 - 15 ngày']['active'])}</td>
            <td>{get_badge_html(stats[am]['Trên 15 ngày']['unattempted'], stats[am]['Trên 15 ngày']['active'])}</td>
            <td style="background-color: #f8fafc; font-weight: 700;">
                {get_badge_html(am_total_unattempted, am_total_active)}
            </td>
        </tr>
        """

    grand_total_row = f"""
    <tr class="total-row">
        <td>TỔNG CỘNG</td>
        <td>{get_badge_html(totals['5 - 8 ngày']['unattempted'], totals['5 - 8 ngày']['active'])}</td>
        <td>{get_badge_html(totals['>8 - 10 ngày']['unattempted'], totals['>8 - 10 ngày']['active'])}</td>
        <td>{get_badge_html(totals['>10 - 15 ngày']['unattempted'], totals['>10 - 15 ngày']['active'])}</td>
        <td>{get_badge_html(totals['Trên 15 ngày']['unattempted'], totals['Trên 15 ngày']['active'])}</td>
        <td style="background-color: #fef08a;">
            {get_badge_html(grand_unattempted, grand_active)}
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
            color: #b91c1c;
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
            border-radius: 6px;
            font-size: 14px;
            font-weight: 700;
            min-width: 40px;
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
            <h2>Báo cáo đơn Aging chưa đi giao lần nào</h2>
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

    temp_html_path = os.path.join(BASE_DIR, "temp_am_unattempted_analysis.html")
    with open(temp_html_path, "w", encoding="utf-8") as f:
        f.write(html_content)

    output_image_path = os.path.join(BASE_DIR, "table_unattempted_aging_analysis.png")
    
    print("   Renderer: Chụp ảnh bằng Playwright...")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.goto(f"file:///{temp_html_path.replace('\\\\', '/')}")
        page.wait_for_timeout(1000)
        container = page.locator("#table-container")
        container.screenshot(path=output_image_path)
        browser.close()

    try:
        os.remove(temp_html_path)
    except:
        pass

    print(f"✔️ Ảnh đã lưu tại: {output_image_path}")

    # 5. Gửi sang GTalk
    print("📡 Đang gửi ảnh báo cáo phân tích chưa giao sang GTalk group...")
    
    caption_text = (
        f"🚨 <b>BÁO CÁO PHÂN TÍCH ĐƠN AGING CHƯA ĐI GIAO LẦN NÀO</b>\n"
        f"📅 Ngày: {datetime.now(tz).strftime('%d-%m-%Y')} | ⏱️ Mốc cập nhật: {datetime.now(tz).strftime('%H:%M')}\n"
        f"========================\n"
        f"{compare_text}\n"
        f"========================\n"
        f"💡 <b>Lưu ý:</b> Ghi nhận việc các AM đã triển khai đi giao các đơn chưa có lượt giao lần nào. Tuy nhiên, đơn hàng cần được xử lý dứt điểm (Giao thành công hoặc Chuyển hoàn thành công) thì mới thực sự thoát khỏi danh sách tồn Aging cao ngày.\n"
        f"========================\n"
        f"🔗 Chi tiết Google Sheet: https://docs.google.com/spreadsheets/d/{SHEET_KEY}/edit#gid=0"
    )
    print("=== NỘI DUNG CAPTION ===")
    print(caption_text)

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
                        resp_send = requests.post("https://mbff.ghn.vn/api/gtalk/send-message", json=send_payload)
                        if resp_send.status_code == 200:
                            print("✅ Đã gửi báo cáo chưa giao lên GTalk group chat thành công!")
                        else:
                            print(f"❌ Lỗi gửi tin nhắn GTalk: {resp_send.status_code} - {resp_send.text}")
    
    # Gửi sang Telegram
    tele_token = "8570130113:AAGXRiUaKBknVpgtm1_i9ZA47JRjAXmB21M"
    tele_chat = "-5058464865"
    print("📡 Đang gửi ảnh báo cáo phân tích chưa giao sang Telegram group...")
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
