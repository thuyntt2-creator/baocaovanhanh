# -*- coding: utf-8 -*-
"""
Script: calculate_report_no_attempt.py
Author: Antigravity AI
Description: Processes the report for "đơn giao > 48h chưa có chuyến giao nào".
Reads data from the 'No attempt' sheet of the Google Sheet, performs AM pivot and snapshot history comparison
against the nearest preceding time slot, writes filtered active orders to individual AM worksheets,
takes a screenshot of the PIVOT table, and posts to GTalk.
"""

import os
import sys
import io
import json
import time
import re
import requests
import asyncio
import pandas as pd
import numpy as np
from datetime import datetime, timedelta, timezone
import gspread
from google.oauth2.service_account import Credentials
from playwright.sync_api import sync_playwright
from dotenv import load_dotenv

# Force UTF-8 output encoding for Task Scheduler / Command Prompt
os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except AttributeError:
    pass

# ============ CONFIGURATION & CONSTANTS ============
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
SNAPSHOT_FILE = os.path.join(BASE_DIR, 'snapshot_no_attempt.json')

SHEET_KEY = '10cq3DUggZ4vXffcxweIRTRK3qiyMeWnV8gksdGwvp7s'
RAW_SHEET_NAME = 'No attempt'
PIVOT_SHEET_NAME = 'PIVOT'

SLOT_COLORS = [
    {"name": "Amber",   "data_bg": "#FEF3C7", "total_bg": "#FDE68A", "header_bg": "#F59E0B", "fg": "#78350F"},
    {"name": "Emerald", "data_bg": "#D1FAE5", "total_bg": "#A7F3D0", "header_bg": "#10B981", "fg": "#065F46"},
    {"name": "Blue",    "data_bg": "#DBEAFE", "total_bg": "#BFDBFE", "header_bg": "#3B82F6", "fg": "#1E40AF"},
    {"name": "Pink",    "data_bg": "#FCE7F3", "total_bg": "#FBCFE8", "header_bg": "#EC4899", "fg": "#9D174D"},
    {"name": "Teal",    "data_bg": "#CCFBF1", "total_bg": "#99F6E4", "header_bg": "#14B8A6", "fg": "#0F766E"},
    {"name": "Purple",  "data_bg": "#F3E5F5", "total_bg": "#E1BEE7", "header_bg": "#8B5CF6", "fg": "#6B21A8"},
    {"name": "Sky",     "data_bg": "#E0F2FE", "total_bg": "#BAE6FD", "header_bg": "#0EA5E9", "fg": "#075985"},
    {"name": "Rose",    "data_bg": "#FFE4E6", "total_bg": "#FECDD3", "header_bg": "#F43F5E", "fg": "#9F1239"}
]

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

GTALK_OA_TOKEN = "2067164759710552066:RfgqBJY4QtV18udu1wMEfhmoRNI4hgBv"
GTALK_CHANNEL_ID = "2067164759710552066"

# Load environment configuration if available
env_path = r"c:\Users\lap4all\Desktop\New folder\.env"
if os.path.exists(env_path):
    load_dotenv(dotenv_path=env_path, override=True)
else:
    load_dotenv()

# Override GTalk credentials from env if defined
GTALK_OA_TOKEN = os.environ.get("GTALK_OA_TOKEN") or GTALK_OA_TOKEN
GTALK_CHANNEL_ID = os.environ.get("NO_ATTEMPT_GTALK_CHANNEL_ID") or GTALK_CHANNEL_ID

# ============ FORMATTING HELPERS ============
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

def format_day_cell(cell_value, cur_val, prev_val, r_idx, c_idx, is_total_row, sheet_id, requests):
    if cur_val is None:
        return ""
    cur_formatted = f"{int(cur_val):,}"
    if prev_val is None:
        cell_text = cur_formatted
        bg_color = "#F9A825" if is_total_row else "#FFF9C4"
        fg_color = "#000000"
        requests.append(cell_format_request(sheet_id, r_idx, r_idx+1, c_idx, c_idx+1, {
            "backgroundColor": make_color(bg_color),
            "textFormat": {"foregroundColor": make_color(fg_color), "bold": True, "fontFamily": "Arial"},
            "horizontalAlignment": "CENTER",
            "verticalAlignment": "MIDDLE",
            "wrapStrategy": "WRAP"
        }))
        return cell_text
        
    delta = cur_val - prev_val
    pct = round(abs(delta) / prev_val * 100) if prev_val != 0 else None
    
    if delta == 0:
        cell_text = f"{cur_formatted}\n(—)"
        bg_color = "#F9A825" if is_total_row else "#FFFFFF"
        fg_color = "#000000"
        bold = True if is_total_row else False
    else:
        pct_str = f" | {pct}%" if pct is not None else ""
        delta_formatted = f"{abs(delta):,}"
        if delta > 0:
            cell_text = f"{cur_formatted}\n(▲{delta_formatted}{pct_str})"
            bg_color = "#FFCDD2"
            fg_color = "#B71C1C"
            bold = True
        else:
            cell_text = f"{cur_formatted}\n(▼{delta_formatted}{pct_str})"
            bg_color = "#C8E6C9"
            fg_color = "#1B5E20"
            bold = True
            
    requests.append(cell_format_request(sheet_id, r_idx, r_idx+1, c_idx, c_idx+1, {
        "backgroundColor": make_color(bg_color),
        "textFormat": {"foregroundColor": make_color(fg_color), "bold": bold, "fontFamily": "Arial"},
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "MIDDLE",
        "wrapStrategy": "WRAP"
    }))
    return cell_text

# ============ TIME PARSER HELPERS ============
def parse_time_to_minutes(t_str):
    """Converts a time string like '07:30' or '10h30' to minutes from midnight."""
    t_str = t_str.replace("h", ":").replace(" ", "")
    m = re.match(r"(\d+):(\d+)", t_str)
    if m:
        return int(m.group(1)) * 60 + int(m.group(2))
    return 0

# ============ MAIN PIPELINE ============
def run_calculations():
    print(f"🔄 Bắt đầu chạy quy trình tính toán lúc: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    
    # 1. Kết nối Google Sheets
    credentials = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(credentials)
    sh = gc_client.open_by_key(SHEET_KEY)
    
    # 2. Đọc dữ liệu từ sheet 'No attempt'
    print(f"📖 Đọc dữ liệu từ sheet '{RAW_SHEET_NAME}'...")
    ws_raw = sh.worksheet(RAW_SHEET_NAME)
    raw_data = ws_raw.get_all_values()
    
    if len(raw_data) < 2:
        print(f"❌ Lỗi: Sheet '{RAW_SHEET_NAME}' không đủ dữ liệu.")
        sys.exit(1)
        
    header = raw_data[0]
    
    # Identify column indices
    try:
        order_col_idx = header.index("ma_don")
        am_col_idx = header.index("am_name")
        status_col_idx = header.index("Trạng thái")
        bc_col_idx = header.index("buu_cuc")
    except ValueError as e:
        print(f"❌ Lỗi: Không tìm thấy cột cần thiết trong sheet '{RAW_SHEET_NAME}'. Chi tiết: {e}")
        sys.exit(1)
        
    # 3. Phân loại trạng thái và đếm số lượng
    # Active orders (Trạng thái khác '#N/A')
    active_orders = []
    
    # Map AM -> {'Chưa gán': 0, 'Đang gán': 0, 'Đã xử lý': 0}
    pivot_map = {}
    
    # Map BC -> {'Chưa gán': 0, 'Đang gán': 0, 'Đã xử lý': 0, 'total': 0, 'am': ''}
    bc_stats = {}
    
    total_chua_gan = 0
    total_dang_gan = 0
    total_na = 0
    
    for row in raw_data[1:]:
        if len(row) <= max(order_col_idx, am_col_idx, status_col_idx, bc_col_idx):
            continue
            
        order_code = row[order_col_idx].strip()
        if not order_code:
            continue
            
        am_name = row[am_col_idx].strip()
        if not am_name:
            am_name = "Không xác định"
            
        status = row[status_col_idx].strip()
        
        # Classification
        if status == '#N/A' or not status:
            category = 'Đã xử lý'
            total_na += 1
        elif status == 'Chưa có chuyến đi trong ngày':
            category = 'Chưa gán'
            total_chua_gan += 1
        else:
            category = 'Đang gán'
            total_dang_gan += 1
            
        # Add to pivot map
        if am_name not in pivot_map:
            pivot_map[am_name] = {'Chưa gán': 0, 'Đang gán': 0, 'Đã xử lý': 0, 'total': 0}
        pivot_map[am_name][category] += 1
        
        # We only count 'Chưa gán' and 'Đang gán' as backlog (total) for comparison and top AMs
        if category in ['Chưa gán', 'Đang gán']:
            pivot_map[am_name]['total'] += 1
            
        # BC Stats
        bc_full_name = row[bc_col_idx].strip()
        if bc_full_name:
            # Extract clean BC name or code
            bc_parts = bc_full_name.split(' - ')
            bc_name = bc_parts[1].strip() if len(bc_parts) > 1 else bc_parts[0].strip()
            
            if bc_name not in bc_stats:
                bc_stats[bc_name] = {'Chưa gán': 0, 'Đang gán': 0, 'Đã xử lý': 0, 'total': 0, 'am': am_name}
            bc_stats[bc_name][category] += 1
            if category in ['Chưa gán', 'Đang gán']:
                bc_stats[bc_name]['total'] += 1
                
        # If the order is active (not resolved / not #N/A), keep it for individual AM worksheets
        if category != 'Đã xử lý':
            active_orders.append(row)
            
    am_names = sorted(list(pivot_map.keys()))
    
    # 4. Quản lý Snapshots
    print("📝 Quản lý Snapshot lịch sử...")
    state = {"last_updated_date": "", "history": [], "daily_snapshots": {}}
    if os.path.exists(SNAPSHOT_FILE):
        try:
            with open(SNAPSHOT_FILE, "r", encoding="utf-8") as f:
                state = json.load(f)
        except:
            pass
            
    # Determine times
    tz = timezone(timedelta(hours=7)) # GMT+7 (Asia/Ho_Chi_Minh)
    now_dt = datetime.now(tz)
    today_str = now_dt.strftime("%Y-%m-%d")
    current_time = now_dt.strftime("%H:%M")
    today_key = now_dt.strftime("%Y%m%d")
    
    # Handle command line custom date/time
    custom_date_str = None
    for idx, arg in enumerate(sys.argv):
        if arg == "--date" and idx + 1 < len(sys.argv):
            custom_date_str = sys.argv[idx + 1]
            break
    if custom_date_str:
        try:
            parsed_dt = datetime.strptime(custom_date_str, "%Y-%m-%d")
            today_str = custom_date_str
            today_key = parsed_dt.strftime("%Y%m%d")
            current_time = "07:30"
            print(f"📅 Sử dụng ngày tùy chỉnh: {today_str} (Mốc: {current_time})")
        except Exception as e:
            print(f"⚠️ Định dạng ngày tùy chỉnh không hợp lệ: {e}")
            
    # Autoclean and recover history from current PIVOT sheet if snapshot file is missing/empty
    if not state.get("daily_snapshots"):
        try:
            print("🔍 Đang phục hồi dữ liệu từ tab PIVOT hiện tại...")
            ws_pivot = sh.worksheet(PIVOT_SHEET_NAME)
            pivot_rows = ws_pivot.get_all_values()
            
            # Locate morning snapshot header row (typically row 30+)
            day_header_row_idx = -1
            for idx, r in enumerate(pivot_rows):
                if len(r) > 0 and "Đơn >2 ngày chưa có lần giao nào — Mốc 7h30 hằng ngày" in r[0]:
                    day_header_row_idx = idx + 1
                    break
            
            if day_header_row_idx != -1 and day_header_row_idx < len(pivot_rows):
                day_header_row = pivot_rows[day_header_row_idx]
                col_dates = {}
                current_year = now_dt.year
                
                # Identify columns with date label (e.g. (23/06))
                for col_idx in range(1, len(day_header_row)):
                    cell_text = day_header_row[col_idx].strip()
                    if not cell_text:
                        continue
                    match = re.search(r'\((\d{2})/(\d{2})\)', cell_text)
                    if match:
                        d_day, d_month = match.groups()
                        date_key = f"{current_year}{d_month}{d_day}"
                        col_dates[col_idx] = date_key
                        
                # Loop AM rows for each date key
                for r_idx in range(day_header_row_idx + 1, len(pivot_rows)):
                    r = pivot_rows[r_idx]
                    if not r or len(r) == 0:
                        continue
                    am_name = r[0].strip()
                    if not am_name or am_name == "TỔNG" or am_name.upper() == "TOTAL":
                        break
                    for col_idx, date_key in col_dates.items():
                        if date_key != today_key and col_idx < len(r):
                            val_str = r[col_idx].strip()
                            if val_str:
                                first_line = val_str.split('\n')[0].strip()
                                num_match = re.match(r'^\d+', first_line)
                                if num_match:
                                    num_val = int(num_match.group())
                                    if "daily_snapshots" not in state:
                                        state["daily_snapshots"] = {}
                                    if date_key not in state["daily_snapshots"]:
                                        state["daily_snapshots"][date_key] = {"totals": {}, "grandTotal": 0}
                                    state["daily_snapshots"][date_key]["totals"][am_name] = num_val
                                    
                # Recalculate daily grand totals
                for dk in state["daily_snapshots"]:
                    totals_map = state["daily_snapshots"][dk]["totals"]
                    state["daily_snapshots"][dk]["grandTotal"] = sum(totals_map.values())
                    
                print("   ✅ Phục hồi lịch sử thành công.")
        except Exception as e:
            print(f"⚠️ Cảnh báo: Không thể phục hồi lịch sử từ PIVOT: {e}")
            
    # Check new day
    if state.get("last_updated_date") != today_str:
        # Archive yesterday's morning run (first run in yesterday's history)
        prev_date = state.get("last_updated_date")
        if prev_date and len(state.get("history", [])) > 0:
            prev_key = prev_date.replace("-", "")
            morning_snap = state["history"][0]
            state["daily_snapshots"][prev_key] = {
                "totals": morning_snap["totals"],
                "bcTotals": morning_snap.get("bcTotals", {}),
                "grandTotal": sum(morning_snap["totals"].values())
            }
        # Clear today's history and start a new day
        state["history"] = []
        state["last_updated_date"] = today_str
        
    # Create current snapshot
    current_am_totals = {am: pivot_map[am]['total'] for am in am_names}
    current_bc_totals = {bc: stats['total'] for bc, stats in bc_stats.items()}
    current_snap = {
        "time": current_time,
        "totals": current_am_totals,
        "bcTotals": current_bc_totals
    }
    
    # Save/update snapshot in history (list of runs for today)
    # Check if time_slot already exists to avoid duplicates
    slot_idx = -1
    for idx, snap in enumerate(state["history"]):
        if snap["time"] == current_time:
            slot_idx = idx
            break
            
    if slot_idx != -1:
        # Overwrite current time slot
        state["history"][slot_idx] = current_snap
    else:
        # Append as new time slot
        state["history"].append(current_snap)
        slot_idx = len(state["history"]) - 1
        
    # Sort history list by time
    state["history"].sort(key=lambda x: parse_time_to_minutes(x["time"]))
    # Recalculate slot index after sorting
    for idx, snap in enumerate(state["history"]):
        if snap["time"] == current_time:
            slot_idx = idx
            break
            
    # Also update today's morning run in daily_snapshots
    state["daily_snapshots"][today_key] = {
        "totals": state["history"][0]["totals"],
        "bcTotals": state["history"][0].get("bcTotals", {}),
        "grandTotal": sum(state["history"][0]["totals"].values())
    }
    
    # Write snapshot state to file
    with open(SNAPSHOT_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)
        
    raw_history = state["history"]
    
    # Determine the comparison snapshot (nearest preceding slot or yesterday's morning run)
    comparison_totals = {}
    comparison_bc_totals = {}
    comparison_time_label = ""
    
    if slot_idx > 0:
        prev_snap = raw_history[slot_idx - 1]
        comparison_totals = prev_snap["totals"]
        comparison_bc_totals = prev_snap["bcTotals"]
        comparison_time_label = f"mốc {prev_snap['time']} hôm nay"
        display_history = [prev_snap, current_snap]
    else:
        # First run of the day, compare with yesterday's morning run
        yesterday_key = (datetime.strptime(today_str, "%Y-%m-%d") - timedelta(days=1)).strftime("%Y%m%d")
        yesterday_snap = state.get("daily_snapshots", {}).get(yesterday_key, {})
        comparison_totals = yesterday_snap.get("totals", {})
        comparison_bc_totals = yesterday_snap.get("bcTotals", {})
        comparison_time_label = "mốc sáng qua (N-1)"
        prev_snap_display = {
            "time": "sáng qua",
            "totals": comparison_totals,
            "bcTotals": comparison_bc_totals
        }
        display_history = [prev_snap_display, current_snap]
        
    history = display_history
    history_len = len(history)
        
    # 5. Ghi dữ liệu và format sheet PIVOT
    print("📊 Cập nhật dữ liệu PIVOT...")
    ws_pivot = sh.worksheet(PIVOT_SHEET_NAME)
    ws_pivot.clear()
    
    # Clear all styling/merges
    clear_format_req = {
        "repeatCell": {
            "range": {
                "sheetId": ws_pivot.id,
                "startRowIndex": 0,
                "endRowIndex": 100,
                "startColumnIndex": 0,
                "endColumnIndex": 20
            },
            "cell": {
                "userEnteredFormat": {}
            },
            "fields": "userEnteredFormat"
        }
    }
    unmerge_req = {
        "unmergeCells": {
            "range": {
                "sheetId": ws_pivot.id,
                "startRowIndex": 0,
                "endRowIndex": 100,
                "startColumnIndex": 0,
                "endColumnIndex": 20
            }
        }
    }
    requests = [clear_format_req, unmerge_req]
    
    # Generate headers for Table 1
    status_cols = ['Chưa gán', 'Đang gán', 'Đã xử lý']
    headers_t1 = ['AM'] + status_cols
    for idx, snap in enumerate(history):
        headers_t1.append(f"Tổng (mốc {snap['time']})")
        if idx > 0:
            headers_t1.append("'+/- so với trước")
            
    grid_values = []
    
    # Table 1: AM Pivot Title
    grid_values.append(['Đơn >2 ngày chưa có lần giao nào'] + [''] * (len(headers_t1) - 1))
    requests.append(merge_request(ws_pivot.id, 0, 1, 0, 4))
    requests.append(merge_request(ws_pivot.id, 0, 1, 4, len(headers_t1)))
    requests.append(cell_format_request(ws_pivot.id, 0, 1, 0, len(headers_t1), {
        "textFormat": {"bold": True, "fontSize": 13, "fontFamily": "Arial"},
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "MIDDLE"
    }))
    requests.append(row_height_request(ws_pivot.id, 0, 1, 35))
    
    # Table 1: Header Row
    grid_values.append(headers_t1)
    requests.append(cell_format_request(ws_pivot.id, 1, 2, 0, len(headers_t1), {
        "backgroundColor": make_color("#5B21B6"),
        "textFormat": {"bold": True, "fontSize": 10, "fontFamily": "Arial", "foregroundColor": make_color("#FFFFFF")},
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "MIDDLE"
    }))
    for idx, snap in enumerate(history):
        color_info = SLOT_COLORS[idx % len(SLOT_COLORS)]
        if idx == 0:
            requests.append(cell_format_request(ws_pivot.id, 1, 2, 4, 5, {
                "backgroundColor": make_color(color_info["header_bg"]),
                "textFormat": {"bold": True, "fontSize": 10, "fontFamily": "Arial", "foregroundColor": make_color("#FFFFFF")},
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE"
            }))
        else:
            col_offset = 3 + idx * 2
            requests.append(cell_format_request(ws_pivot.id, 1, 2, col_offset, col_offset + 2, {
                "backgroundColor": make_color(color_info["header_bg"]),
                "textFormat": {"bold": True, "fontSize": 10, "fontFamily": "Arial", "foregroundColor": make_color("#FFFFFF")},
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE"
            }))
    requests.append(row_height_request(ws_pivot.id, 1, 2, 30))
    
    # Table 1: AM Rows (sorted by total backlog descending)
    am_rows_start = 2
    sorted_ams = sorted(am_names, key=lambda x: pivot_map[x]['total'], reverse=True)
    
    for r_idx, am in enumerate(sorted_ams):
        row_idx = am_rows_start + r_idx
        row = [am]
        for col in status_cols:
            row.append(pivot_map[am][col])
            
        for idx, snap in enumerate(history):
            snap_val = snap["totals"].get(am, 0)
            row.append(snap_val)
            if idx > 0:
                prev_val = history[idx - 1]["totals"].get(am, 0)
                diff = snap_val - prev_val
                row.append("Không đổi" if diff == 0 else f"+ {diff}" if diff > 0 else f"- {abs(diff)}")
        grid_values.append(row)
        
        bg_color = "#F5F3FF" if r_idx % 2 == 0 else "#FFFFFF"
        requests.append(cell_format_request(ws_pivot.id, row_idx, row_idx+1, 0, 1, {
            "backgroundColor": make_color("#EDE9FE"),
            "textFormat": {"bold": True, "fontFamily": "Arial", "fontSize": 9},
            "horizontalAlignment": "LEFT",
            "verticalAlignment": "MIDDLE"
        }))
        requests.append(cell_format_request(ws_pivot.id, row_idx, row_idx+1, 1, 4, {
            "backgroundColor": make_color(bg_color),
            "textFormat": {"fontFamily": "Arial", "fontSize": 9},
            "horizontalAlignment": "CENTER",
            "verticalAlignment": "MIDDLE",
            "numberFormat": {"type": "NUMBER", "pattern": "#,##0"}
        }))
        # Highlight total columns
        color_info = SLOT_COLORS[0]
        requests.append(cell_format_request(ws_pivot.id, row_idx, row_idx+1, 4, 5, {
            "backgroundColor": make_color(color_info["data_bg"]),
            "textFormat": {"bold": True, "fontFamily": "Arial", "fontSize": 9},
            "horizontalAlignment": "CENTER",
            "verticalAlignment": "MIDDLE",
            "numberFormat": {"type": "NUMBER", "pattern": "#,##0"}
        }))
        
        for idx in range(1, history_len):
            col_offset = 3 + idx * 2
            color_info = SLOT_COLORS[idx % len(SLOT_COLORS)]
            # Total slot col
            requests.append(cell_format_request(ws_pivot.id, row_idx, row_idx+1, col_offset, col_offset+1, {
                "backgroundColor": make_color(color_info["data_bg"]),
                "textFormat": {"bold": True, "fontFamily": "Arial", "fontSize": 9},
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE",
                "numberFormat": {"type": "NUMBER", "pattern": "#,##0"}
            }))
            # +/- col
            diff_val = row[col_offset+1]
            d_bg, d_fg = ("#FFCDD2", "#B71C1C") if diff_val.startswith('+') else ("#C8E6C9", "#1B5E20") if diff_val.startswith('-') else (color_info["data_bg"], "#000000")
            requests.append(cell_format_request(ws_pivot.id, row_idx, row_idx+1, col_offset+1, col_offset+2, {
                "backgroundColor": make_color(d_bg),
                "textFormat": {"bold": True, "fontFamily": "Arial", "fontSize": 9, "foregroundColor": make_color(d_fg)},
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE"
            }))
        requests.append(row_height_request(ws_pivot.id, row_idx, row_idx+1, 24))
        
    # Table 1: Total Row
    total_row_idx = am_rows_start + len(sorted_ams)
    total_row_val = ['TỔNG']
    for col in status_cols:
        total_row_val.append(sum(pivot_map[am][col] for am in am_names))
        
    for idx, snap in enumerate(history):
        t_sum = sum(snap["totals"].get(am, 0) for am in am_names)
        total_row_val.append(t_sum)
        if idx > 0:
            prev_t_sum = sum(history[idx - 1]["totals"].get(am, 0) for am in am_names)
            t_diff = t_sum - prev_t_sum
            total_row_val.append("Không đổi" if t_diff == 0 else f"+ {t_diff}" if t_diff > 0 else f"- {abs(t_diff)}")
    grid_values.append(total_row_val)
    
    requests.append(cell_format_request(ws_pivot.id, total_row_idx, total_row_idx+1, 0, len(headers_t1), {
        "backgroundColor": make_color("#FFF176"),
        "textFormat": {"bold": True, "fontFamily": "Arial", "fontSize": 10},
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "MIDDLE"
    }))
    requests.append(cell_format_request(ws_pivot.id, total_row_idx, total_row_idx+1, 0, 1, {
        "horizontalAlignment": "LEFT"
    }))
    color_info = SLOT_COLORS[0]
    requests.append(cell_format_request(ws_pivot.id, total_row_idx, total_row_idx+1, 4, 5, {
        "backgroundColor": make_color(color_info["total_bg"])
    }))
    
    for idx in range(1, history_len):
        col_offset = 3 + idx * 2
        color_info = SLOT_COLORS[idx % len(SLOT_COLORS)]
        requests.append(cell_format_request(ws_pivot.id, total_row_idx, total_row_idx+1, col_offset, col_offset+1, {
            "backgroundColor": make_color(color_info["total_bg"])
        }))
        t_diff_val = total_row_val[col_offset+1]
        td_bg, td_fg = ("#FFCDD2", "#B71C1C") if t_diff_val.startswith('+') else ("#C8E6C9", "#1B5E20") if t_diff_val.startswith('-') else (color_info["total_bg"], "#000000")
        requests.append(cell_format_request(ws_pivot.id, total_row_idx, total_row_idx+1, col_offset+1, col_offset+2, {
            "backgroundColor": make_color(td_bg),
            "textFormat": {"bold": True, "fontFamily": "Arial", "fontSize": 10, "foregroundColor": make_color(td_fg)},
            "horizontalAlignment": "CENTER",
            "verticalAlignment": "MIDDLE"
        }))
    requests.append(row_height_request(ws_pivot.id, total_row_idx, total_row_idx+1, 24))
    requests.append(border_request(ws_pivot.id, 1, total_row_idx + 1, 0, len(headers_t1)))
    
    # Spacer row
    grid_values.append([''] * len(headers_t1))
    
    # Table 2: Top 5 BCs
    bc_start_row_idx = total_row_idx + 2
    grid_values.append(['Top 5 BC tồn nhiều nhất'] + [''] * (len(headers_t1) - 1))
    requests.append(merge_request(ws_pivot.id, bc_start_row_idx, bc_start_row_idx+1, 0, 4))
    requests.append(merge_request(ws_pivot.id, bc_start_row_idx, bc_start_row_idx+1, 4, len(headers_t1)))
    requests.append(cell_format_request(ws_pivot.id, bc_start_row_idx, bc_start_row_idx+1, 0, len(headers_t1), {
        "textFormat": {"bold": True, "fontSize": 13, "fontFamily": "Arial"},
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "MIDDLE"
    }))
    requests.append(row_height_request(ws_pivot.id, bc_start_row_idx, bc_start_row_idx+1, 35))
    
    bc_headers = ['Bưu cục'] + status_cols
    for idx, snap in enumerate(history):
        bc_headers.append(f"Tổng (mốc {snap['time']})")
        if idx > 0:
            bc_headers.append("'+/- so với trước")
    grid_values.append(bc_headers)
    requests.append(cell_format_request(ws_pivot.id, bc_start_row_idx+1, bc_start_row_idx+2, 0, len(headers_t1), {
        "backgroundColor": make_color("#5B21B6"),
        "textFormat": {"bold": True, "fontSize": 10, "fontFamily": "Arial", "foregroundColor": make_color("#FFFFFF")},
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "MIDDLE"
    }))
    for idx, snap in enumerate(history):
        color_info = SLOT_COLORS[idx % len(SLOT_COLORS)]
        if idx == 0:
            requests.append(cell_format_request(ws_pivot.id, bc_start_row_idx+1, bc_start_row_idx+2, 4, 5, {
                "backgroundColor": make_color(color_info["header_bg"]),
                "textFormat": {"bold": True, "fontSize": 10, "fontFamily": "Arial", "foregroundColor": make_color("#FFFFFF")},
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE"
            }))
        else:
            col_offset = 3 + idx * 2
            requests.append(cell_format_request(ws_pivot.id, bc_start_row_idx+1, bc_start_row_idx+2, col_offset, col_offset + 2, {
                "backgroundColor": make_color(color_info["header_bg"]),
                "textFormat": {"bold": True, "fontSize": 10, "fontFamily": "Arial", "foregroundColor": make_color("#FFFFFF")},
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE"
            }))
    requests.append(row_height_request(ws_pivot.id, bc_start_row_idx+1, bc_start_row_idx+2, 30))
    
    sorted_bcs = sorted(bc_stats.items(), key=lambda x: x[1]['total'], reverse=True)
    top5_bc_names = [x[0] for x in sorted_bcs[:5]]
    
    for r_idx, bc_name in enumerate(top5_bc_names):
        row_idx = bc_start_row_idx + 2 + r_idx
        row = [bc_name]
        for col in status_cols:
            row.append(bc_stats[bc_name].get(col, 0))
            
        for idx, snap in enumerate(history):
            snap_val = snap["bcTotals"].get(bc_name, 0)
            row.append(snap_val)
            if idx > 0:
                prev_val = history[idx - 1]["bcTotals"].get(bc_name, 0)
                diff = snap_val - prev_val
                row.append("Không đổi" if diff == 0 else f"+ {diff}" if diff > 0 else f"- {abs(diff)}")
        grid_values.append(row)
        
        bg_color = "#F5F3FF" if r_idx % 2 == 0 else "#FFFFFF"
        requests.append(cell_format_request(ws_pivot.id, row_idx, row_idx+1, 0, 1, {
            "backgroundColor": make_color("#EDE9FE"),
            "textFormat": {"bold": True, "fontFamily": "Arial", "fontSize": 9},
            "horizontalAlignment": "LEFT",
            "verticalAlignment": "MIDDLE"
        }))
        requests.append(cell_format_request(ws_pivot.id, row_idx, row_idx+1, 1, 4, {
            "backgroundColor": make_color(bg_color),
            "textFormat": {"fontFamily": "Arial", "fontSize": 9},
            "horizontalAlignment": "CENTER",
            "verticalAlignment": "MIDDLE",
            "numberFormat": {"type": "NUMBER", "pattern": "#,##0"}
        }))
        color_info = SLOT_COLORS[0]
        requests.append(cell_format_request(ws_pivot.id, row_idx, row_idx+1, 4, 5, {
            "backgroundColor": make_color(color_info["data_bg"]),
            "textFormat": {"bold": True, "fontFamily": "Arial", "fontSize": 9},
            "horizontalAlignment": "CENTER",
            "verticalAlignment": "MIDDLE",
            "numberFormat": {"type": "NUMBER", "pattern": "#,##0"}
        }))
        
        for idx in range(1, history_len):
            col_offset = 3 + idx * 2
            color_info = SLOT_COLORS[idx % len(SLOT_COLORS)]
            requests.append(cell_format_request(ws_pivot.id, row_idx, row_idx+1, col_offset, col_offset+1, {
                "backgroundColor": make_color(color_info["data_bg"]),
                "textFormat": {"bold": True, "fontFamily": "Arial", "fontSize": 9},
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE",
                "numberFormat": {"type": "NUMBER", "pattern": "#,##0"}
            }))
            diff_val = row[col_offset+1]
            d_bg, d_fg = ("#FFCDD2", "#B71C1C") if diff_val.startswith('+') else ("#C8E6C9", "#1B5E20") if diff_val.startswith('-') else (color_info["data_bg"], "#000000")
            requests.append(cell_format_request(ws_pivot.id, row_idx, row_idx+1, col_offset+1, col_offset+2, {
                "backgroundColor": make_color(d_bg),
                "textFormat": {"bold": True, "fontFamily": "Arial", "fontSize": 9, "foregroundColor": make_color(d_fg)},
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE"
            }))
        requests.append(row_height_request(ws_pivot.id, row_idx, row_idx+1, 24))
        
    bc_end_row_idx = bc_start_row_idx + 2 + len(top5_bc_names)
    requests.append(border_request(ws_pivot.id, bc_start_row_idx + 1, bc_end_row_idx, 0, len(headers_t1)))
    
    # Spacer rows
    grid_values.append([''] * len(headers_t1))
    grid_values.append([''] * len(headers_t1))
    
    # Table 3: 8-Day Morning Snapshot History
    day_start_row_idx = bc_end_row_idx + 2
    grid_values.append(['Đơn >2 ngày chưa có lần giao nào — Mốc 7h30 hằng ngày'] + [''] * 8)
    requests.append(merge_request(ws_pivot.id, day_start_row_idx, day_start_row_idx+1, 0, 9))
    requests.append(cell_format_request(ws_pivot.id, day_start_row_idx, day_start_row_idx+1, 0, 9, {
        "textFormat": {"bold": True, "fontSize": 13, "fontFamily": "Arial"},
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "MIDDLE"
    }))
    requests.append(row_height_request(ws_pivot.id, day_start_row_idx, day_start_row_idx+1, 35))
    
    # Daily snapshots labels (N to N-7)
    anchor_dt = datetime.strptime(today_str, "%Y-%m-%d")
    date_keys = []
    date_labels = []
    for d in range(8):
        dt = anchor_dt - timedelta(days=d)
        date_keys.append(dt.strftime("%Y%m%d"))
        date_labels.append(f"Ngày N-{d}\n({dt.strftime('%d/%m')})" if d > 0 else f"Ngày N\n({dt.strftime('%d/%m')})")
        
    grid_values.append(['AM'] + date_labels)
    requests.append(cell_format_request(ws_pivot.id, day_start_row_idx+1, day_start_row_idx+2, 0, 9, {
        "backgroundColor": make_color("#5B21B6"),
        "textFormat": {"bold": True, "fontSize": 10, "fontFamily": "Arial", "foregroundColor": make_color("#FFFFFF")},
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "MIDDLE",
        "wrapStrategy": "WRAP"
    }))
    requests.append(row_height_request(ws_pivot.id, day_start_row_idx+1, day_start_row_idx+2, 40))
    
    daily_snapshots = state.get("daily_snapshots", {})
    day_am_start_idx = day_start_row_idx + 2
    
    for r_idx, am in enumerate(sorted_ams):
        row_idx = day_am_start_idx + r_idx
        row = [am]
        requests.append(cell_format_request(ws_pivot.id, row_idx, row_idx+1, 0, 1, {
            "backgroundColor": make_color("#EDE9FE"),
            "textFormat": {"bold": True, "fontFamily": "Arial", "fontSize": 9},
            "horizontalAlignment": "LEFT",
            "verticalAlignment": "MIDDLE"
        }))
        
        for d in range(8):
            dk = date_keys[d]
            cur_val = daily_snapshots.get(dk, {}).get("totals", {}).get(am, None)
            dk_prev = (datetime.strptime(dk, "%Y%m%d") - timedelta(days=1)).strftime("%Y%m%d")
            prev_val = daily_snapshots.get(dk_prev, {}).get("totals", {}).get(am, None)
            
            cell_text = format_day_cell(None, cur_val, prev_val, row_idx, d+1, False, ws_pivot.id, requests)
            row.append(cell_text)
            
        grid_values.append(row)
        requests.append(row_height_request(ws_pivot.id, row_idx, row_idx+1, 40))
        
    # Total row daily snapshots
    day_total_row_idx = day_am_start_idx + len(sorted_ams)
    day_total_row = ['TỔNG']
    requests.append(cell_format_request(ws_pivot.id, day_total_row_idx, day_total_row_idx+1, 0, 1, {
        "backgroundColor": make_color("#FFF176"),
        "textFormat": {"bold": True, "fontFamily": "Arial", "fontSize": 10},
        "horizontalAlignment": "LEFT",
        "verticalAlignment": "MIDDLE"
    }))
    
    for d in range(8):
        dk = date_keys[d]
        cur_tot = daily_snapshots.get(dk, {}).get("grandTotal", None)
        dk_prev = (datetime.strptime(dk, "%Y%m%d") - timedelta(days=1)).strftime("%Y%m%d")
        prev_tot = daily_snapshots.get(dk_prev, {}).get("grandTotal", None)
        
        cell_text = format_day_cell(None, cur_tot, prev_tot, day_total_row_idx, d+1, True, ws_pivot.id, requests)
        day_total_row.append(cell_text)
        
    grid_values.append(day_total_row)
    requests.append(row_height_request(ws_pivot.id, day_total_row_idx, day_total_row_idx+1, 40))
    requests.append(border_request(ws_pivot.id, day_start_row_idx+1, day_total_row_idx+1, 0, 9))
    
    # Set column widths specifically for PIVOT
    col_widths = {0: 380, 1: 100, 2: 100, 3: 100, 4: 120, 5: 120, 6: 120, 7: 120, 8: 120, 9: 120}
    for c_idx, w in col_widths.items():
        requests.append(col_width_request(ws_pivot.id, c_idx, c_idx+1, w))
        
    # Write PIVOT grid data
    max_cols_val = max(len(r) for r in grid_values)
    clean_grid = []
    for r in grid_values:
        if len(r) < max_cols_val:
            r = r + [''] * (max_cols_val - len(r))
        clean_grid.append(r)
        
    end_col_letter = gspread.utils.rowcol_to_a1(1, len(clean_grid[0])).split("1")[0]
    ws_pivot.update(range_name=f"A1:{end_col_letter}{len(clean_grid)}", values=clean_grid, value_input_option="USER_ENTERED")
    sh.batch_update({"requests": requests})
    print("✔️ Đã cập nhật xong sheet 'PIVOT'.")
    
    # 6. Tách đơn theo AM và ghi vào worksheets AM
    print("📂 Đang tách đơn theo AM...")
    am_groups = {}
    for o in active_orders:
        am = o[am_col_idx].strip()
        if not am:
            am = "Không xác định"
        if am not in am_groups:
            am_groups[am] = []
        am_groups[am].append(o)
        
    all_worksheets = {ws.title: ws for ws in sh.worksheets()}
    am_links = {}
    
    for am_name in am_names:
        am_rows = am_groups.get(am_name, [])
        if am_name in all_worksheets:
            ws_am = all_worksheets[am_name]
            ws_am.clear()
        else:
            ws_am = sh.add_worksheet(title=am_name, rows=str(max(100, len(am_rows) + 50)), cols="15")
            
        ws_am.update([header] + am_rows)
        am_links[am_name] = f"https://docs.google.com/spreadsheets/d/{SHEET_KEY}/edit#gid={ws_am.id}"
        
        # Format AM worksheets headers
        sh.batch_update({
            "requests": [
                cell_format_request(ws_am.id, 0, 1, 0, 12, {
                    "backgroundColor": make_color("#5B21B6"),
                    "textFormat": {"bold": True, "fontSize": 10, "fontFamily": "Arial", "foregroundColor": make_color("#FFFFFF")},
                    "horizontalAlignment": "CENTER",
                    "verticalAlignment": "MIDDLE"
                }),
                row_height_request(ws_am.id, 0, 1, 28)
            ]
        })
        
    print("✔️ Đã cập nhật xong các sheet AM cá nhân.")
    
    # 7. Render HTML & Playwright Chụp ảnh PIVOT & Gửi GTalk
    print("📸 Khởi chạy Playwright để chụp hình bảng PIVOT...")
    try:
        generate_and_send_colored_table(pivot_map, am_names, status_cols, history, top5_bc_names, bc_stats, current_time, today_str, am_links, state, comparison_totals, comparison_bc_totals, comparison_time_label, ws_pivot.id)
    except Exception as e:
        print(f"❌ Lỗi khi render/gửi báo cáo: {e}")
        import traceback
        traceback.print_exc()

def generate_and_send_colored_table(pivot_map, am_names, status_cols, history, top5_bc_names, bc_stats, current_time, today_str, am_links, state, comparison_totals, comparison_bc_totals, comparison_time_label, ws_pivot_id):
    history_len = len(history)
    
    # Build AM table rows
    body_rows_am = ""
    sorted_ams = sorted(am_names, key=lambda x: pivot_map[x]['total'], reverse=True)
    for idx, am in enumerate(sorted_ams):
        chua_gan = pivot_map[am]['Chưa gán']
        dang_gan = pivot_map[am]['Đang gán']
        na_val = pivot_map[am]['Đã xử lý']
        
        history_cols_html = ""
        for snap_idx, snap in enumerate(history):
            color_info = SLOT_COLORS[snap_idx % len(SLOT_COLORS)]
            d_bg = color_info["data_bg"]
            snap_val = snap["totals"].get(am, 0)
            history_cols_html += f'<td style="background-color: {d_bg} !important;">{snap_val:,}</td>'
            if snap_idx > 0:
                prev_val = history[snap_idx - 1]["totals"].get(am, 0)
                diff = snap_val - prev_val
                if diff > 0:
                    badge = f'<span class="delta-badge delta-red">▲ +{diff}</span>'
                    td_bg = '#FFCDD2'
                elif diff < 0:
                    badge = f'<span class="delta-badge delta-green">▼ {abs(diff)}</span>'
                    td_bg = '#C8E6C9'
                else:
                    badge = '<span class="delta-none">—</span>'
                    td_bg = d_bg
                history_cols_html += f'<td style="background-color: {td_bg} !important;">{badge}</td>'
                
        am_link_html = f'<a href="{am_links.get(am, "#")}" target="_blank" style="text-decoration:none; color:#6D28D9; font-weight:bold;">{am}</a>'
        body_rows_am += f"""
        <tr>
            <td class="left-align">{am_link_html}</td>
            <td>{chua_gan:,}</td>
            <td>{dang_gan:,}</td>
            <td>{na_val:,}</td>
            {history_cols_html}
        </tr>
        """
        
    # Totals Row AM table
    total_chua_gan = sum(pivot_map[am]['Chưa gán'] for am in am_names)
    total_dang_gan = sum(pivot_map[am]['Đang gán'] for am in am_names)
    total_na = sum(pivot_map[am]['Đã xử lý'] for am in am_names)
    
    total_history_html = ""
    for snap_idx, snap in enumerate(history):
        color_info = SLOT_COLORS[snap_idx % len(SLOT_COLORS)]
        t_bg = color_info["total_bg"]
        snap_sum = sum(snap["totals"].get(am, 0) for am in am_names)
        total_history_html += f'<td style="background-color: {t_bg} !important; color: #000000; font-weight: 800;">{snap_sum:,}</td>'
        if snap_idx > 0:
            prev_sum = sum(history[snap_idx - 1]["totals"].get(am, 0) for am in am_names)
            diff = snap_sum - prev_sum
            if diff > 0:
                badge = f'<span class="delta-badge delta-red">▲ +{diff}</span>'
                td_bg = '#FFCDD2'
            elif diff < 0:
                badge = f'<span class="delta-badge delta-green">▼ {abs(diff)}</span>'
                td_bg = '#C8E6C9'
            else:
                badge = '<span class="delta-none">—</span>'
                td_bg = t_bg
            total_history_html += f'<td style="background-color: {td_bg} !important;">{badge}</td>'
            
    am_total_row_html = f"""
    <tr class="grand-total">
        <td class="left-align">TỔNG CỘNG</td>
        <td>{total_chua_gan:,}</td>
        <td>{total_dang_gan:,}</td>
        <td>{total_na:,}</td>
        {total_history_html}
    </tr>
    """
    
    # Columns for Table 2 (Top 5 BCs)
    body_rows_bc = ""
    for idx, bc_name in enumerate(top5_bc_names):
        chua_gan = bc_stats[bc_name].get('Chưa gán', 0)
        dang_gan = bc_stats[bc_name].get('Đang gán', 0)
        na_val = bc_stats[bc_name].get('Đã xử lý', 0)
        
        history_cols_html = ""
        for snap_idx, snap in enumerate(history):
            color_info = SLOT_COLORS[snap_idx % len(SLOT_COLORS)]
            d_bg = color_info["data_bg"]
            snap_val = snap["bcTotals"].get(bc_name, 0)
            history_cols_html += f'<td style="background-color: {d_bg} !important;">{snap_val:,}</td>'
            if snap_idx > 0:
                prev_val = history[snap_idx - 1]["bcTotals"].get(bc_name, 0)
                diff = snap_val - prev_val
                if diff > 0:
                    badge = f'<span class="delta-badge delta-red">▲ +{diff}</span>'
                    td_bg = '#FFCDD2'
                elif diff < 0:
                    badge = f'<span class="delta-badge delta-green">▼ {abs(diff)}</span>'
                    td_bg = '#C8E6C9'
                else:
                    badge = '<span class="delta-none">—</span>'
                    td_bg = d_bg
                history_cols_html += f'<td style="background-color: {td_bg} !important;">{badge}</td>'
                
        body_rows_bc += f"""
        <tr>
            <td class="left-align" style="font-weight:bold;">{bc_name}</td>
            <td>{chua_gan:,}</td>
            <td>{dang_gan:,}</td>
            <td>{na_val:,}</td>
            {history_cols_html}
        </tr>
        """
        
    # Table 1 & 2 Headers columns count span
    headers_snap_html = ""
    subheaders_snap_html = ""
    for snap_idx, snap in enumerate(history):
        color_info = SLOT_COLORS[snap_idx % len(SLOT_COLORS)]
        h_bg = color_info["header_bg"]
        headers_snap_html += f'<th rowspan="2" style="background: {h_bg} !important;">Tổng<br>(mốc {snap["time"]})</th>'
        if snap_idx > 0:
            headers_snap_html += f'<th rowspan="2" style="background: {h_bg} !important;">+/-</th>'
            
    # Build complete HTML string for elements screenshots
    html_content = f"""<!DOCTYPE html>
    <html>
    <head>
    <meta charset="utf-8">
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <style>
      body {{
        font-family: 'Inter', sans-serif;
        margin: 0;
        padding: 20px;
        background-color: #FFFFFF;
        display: flex;
        flex-direction: column;
        align-items: center;
        gap: 30px;
      }}
      #capture-container {{
        background: #ffffff;
        padding: 36px;
        border-radius: 20px;
        box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.1), 0 10px 10px -5px rgba(0, 0, 0, 0.04);
        border: 1px solid rgba(255, 255, 255, 0.8);
        max-width: 1700px;
        width: 100%;
        box-sizing: border-box;
      }}
      .tables-wrapper {{
        display: flex;
        gap: 30px;
        align-items: flex-start;
        width: 100%;
      }}
      .table-section {{
        flex: 1;
        min-width: 0;
      }}
      .header {{
        display: flex;
        justify-content: space-between;
        align-items: center;
        border-bottom: 2px solid #e2e8f0;
        padding-bottom: 20px;
        margin-bottom: 25px;
        width: 100%;
      }}
      .header h2 {{
        margin: 0;
        font-family: 'Inter', sans-serif;
        font-size: 28px;
        background: linear-gradient(90deg, #4C1D95 0%, #7C3AED 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-weight: 800;
        text-transform: uppercase;
        letter-spacing: 0.5px;
      }}
      .header .time {{
        font-size: 14px;
        color: #475569;
        font-weight: 600;
        background: #f1f5f9;
        padding: 6px 14px;
        border-radius: 30px;
        border: 1px solid #e2e8f0;
      }}
      .table-title {{
        font-family: 'Inter', sans-serif;
        font-size: 22px;
        font-weight: 800;
        color: #4C1D95;
        margin-top: 10px;
        margin-bottom: 16px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        border-left: 5px solid #7C3AED;
        padding-left: 12px;
        text-align: left;
      }}
      table {{
        width: 100%;
        border-collapse: separate;
        border-spacing: 0;
        text-align: left;
        margin-bottom: 35px;
        border-radius: 12px;
        overflow: hidden;
        border: 1px solid #e2e8f0;
      }}
      th {{
        font-family: 'Inter', sans-serif;
        background: linear-gradient(180deg, #5B21B6 0%, #4C1D95 100%);
        color: #ffffff;
        font-weight: 750;
        font-size: 14px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        padding: 12px 10px;
        text-align: center;
        border: none;
      }}
      th.left-align, td.left-align {{
        text-align: left;
        padding-left: 20px;
      }}
      td {{
        padding: 11px 9px;
        font-size: 13px;
        color: #334155;
        border-bottom: 1px solid #f1f5f9;
        font-weight: 600;
        text-align: center;
        background-color: #ffffff;
      }}
      tr:nth-child(even) td {{
        background-color: #f5f3ff;
      }}
      tr:hover td {{
        background-color: #ede9fe;
      }}
      .grand-total td {{
        background: #fef08a !important;
        color: #854d0e;
        font-weight: 800;
        border-top: 2px solid #eab308;
        border-bottom: none;
      }}
      .delta-badge {{
        display: inline-block;
        padding: 2px 6px;
        border-radius: 4px;
        font-size: 11px;
        font-weight: bold;
      }}
      .delta-red {{
        color: #B71C1C;
        background-color: #FFCDD2;
      }}
      .delta-green {{
        color: #1B5E20;
        background-color: #C8E6C9;
      }}
      .delta-none {{
        color: #7f8c8d;
      }}
    </style>
    </head>
    <body>
      <div id="capture-container">
        <div class="header">
          <h2>Báo cáo đơn giao > 48h chưa có chuyến giao nào</h2>
          <div class="time">Mốc cập nhật: {current_time} ngày {today_str}</div>
        </div>
        
        <div class="tables-wrapper">
          <!-- TABLE 1: AM PIVOT -->
          <div class="table-section">
            <div class="table-title">Thống kê theo AM</div>
            <table>
              <thead>
                <tr>
                  <th rowspan="2" class="left-align">AM</th>
                  <th colspan="3" style="background-color: #7C3AED;">Trạng thái chi tiết</th>
                  {headers_snap_html}
                </tr>
                <tr>
                  <th style="background-color: #7C3AED;">Chưa gán</th>
                  <th style="background-color: #7C3AED;">Đang gán</th>
                  <th style="background-color: #7C3AED;">Đã xử lý</th>
                </tr>
              </thead>
              <tbody>
                {body_rows_am}
                {am_total_row_html}
              </tbody>
            </table>
          </div>
          
          <!-- TABLE 2: TOP 5 BC -->
          <div class="table-section">
            <div class="table-title">Top 5 Bưu cục tồn chưa giao cao nhất</div>
            <table>
              <thead>
                <tr>
                  <th rowspan="2" class="left-align">Bưu cục</th>
                  <th colspan="3" style="background-color: #7C3AED;">Trạng thái chi tiết</th>
                  {headers_snap_html}
                </tr>
                <tr>
                  <th style="background-color: #7C3AED;">Chưa gán</th>
                  <th style="background-color: #7C3AED;">Đang gán</th>
                  <th style="background-color: #7C3AED;">Đã xử lý</th>
                </tr>
              </thead>
              <tbody>
                {body_rows_bc}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </body>
    </html>
    """
    
    temp_html_path = os.path.join(BASE_DIR, 'temp_no_attempt.html')
    output_image_path = os.path.join(BASE_DIR, 'table_no_attempt.png')
    
    with open(temp_html_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
        
    # Screenshot using Playwright
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page()
        page.set_viewport_size({"width": 1750, "height": 1200})
        page.goto(f"file:///{os.path.abspath(temp_html_path)}")
        page.wait_for_timeout(500)
        
        # Take screenshot of the tables together
        element = page.locator("#capture-container")
        element.screenshot(path=output_image_path)
        browser.close()
        
    try:
        os.remove(temp_html_path)
    except:
        pass
        
    # 8. Send Gtalk Notification
    print("📡 Đang gửi ảnh và báo cáo sang GTalk group...")
    file_name = os.path.basename(output_image_path)
    file_size = os.path.getsize(output_image_path)
    
    with open(output_image_path, 'rb') as f:
        file_bytes = f.read()
        
    sheet_link = f"https://docs.google.com/spreadsheets/d/{SHEET_KEY}/edit#gid={ws_pivot_id}"
    
    # Calculate Gtalk text summary
    caption = f"📊 <b>Báo cáo Đơn giao > 48h chưa có chuyến giao nào {today_str}</b>\n"
    caption += f"⏱️ <b>Mốc cập nhật:</b> {current_time}\n"
    caption += f"🔗 <b>Xem chi tiết danh sách AM:</b> <a href=\"{sheet_link}\"><b>Link báo cáo</b></a>\n\n"
    
    caption += f"🏆 <b>TOP AM TỒN CHƯA GIAO CAO NHẤT (so với {comparison_time_label}):</b>\n"
    sorted_ams_by_backlog = sorted(am_names, key=lambda x: pivot_map[x]['total'], reverse=True)
    for idx, am in enumerate(sorted_ams_by_backlog[:5]):
        current_val = pivot_map[am]['total']
        prev_val = comparison_totals.get(am, None)
        
        change_text = ""
        if prev_val is not None:
            diff = current_val - prev_val
            pct_str = ""
            if prev_val > 0:
                pct = round(abs(diff) / prev_val * 100)
                pct_str = f" ~ giảm {pct}%" if diff < 0 else f" ~ tăng {pct}%" if diff > 0 else ""
                
            if diff < 0:
                change_text = f" (Giảm {abs(diff)} đơn{pct_str})"
            elif diff > 0:
                change_text = f" (Tồn tăng thêm +{diff} đơn{pct_str})"
            else:
                change_text = " (Không đổi)"
        else:
            change_text = ""
            
        caption += f"  {idx+1}. AM <b>{am}</b>: <b>{current_val}</b> đơn{change_text}\n"
        

    
    init_payload = {
        "ChannelId": GTALK_CHANNEL_ID,
        "FileName": file_name,
        "FileSize": str(file_size),
        "MimeType": "image/png",
        "Metadata": json.dumps({"width": 1700, "height": 1000}),
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
                            "clientMsgId": str(int(os.path.getmtime(output_image_path) * 1000)),
                            "content": {
                                "parseMode": "HTML",
                                "attachment": {
                                    "caption": caption,
                                    "items": [{"image": {"fileId": file_id, "width": 1700, "height": 1000}}]
                                }
                            },
                            "oaToken": GTALK_OA_TOKEN
                        }
                        r_send = requests.post("https://mbff.ghn.vn/api/gtalk/send-message", json=send_payload)
                        if r_send.status_code == 200 and r_send.json().get("errorCode") == "success":
                            print("   ✅ Đã gửi báo cáo sang Gtalk group thành công!")
                        else:
                            print(f"   ❌ Gửi tin nhắn GTalk lỗi: {r_send.text}")
                            
    # Clean up local image after sending
    try:
        os.remove(output_image_path)
    except:
        pass

def main():
    current_hour = datetime.now().hour
    bypass_time = len(sys.argv) > 1 and sys.argv[1] == "--force"
    if not bypass_time and not (7 <= current_hour <= 22):
        print(f"💤 Ngoài khung giờ hoạt động (7h - 22h). Hiện tại là {datetime.now().strftime('%H:%M:%S')}. Script sẽ dừng.")
        print("💡 Để chạy bất chấp khung giờ này, vui lòng thêm tham số --force khi chạy (Ví dụ: CHAY_BAO_CAO_NO_ATTEMPT.bat --force)")
        sys.exit(2)
        
    # Clear snapshot history if --clear argument is passed
    if "--clear" in sys.argv:
        print("🧹 Đang xóa lịch sử...")
        if os.path.exists(SNAPSHOT_FILE):
            try:
                os.remove(SNAPSHOT_FILE)
                print("✔️ Đã xóa lịch sử hôm nay thành công.")
            except Exception as e:
                print(f"❌ Lỗi khi xóa lịch sử: {e}")
                
    try:
        run_calculations()
    except Exception as e:
        print(f"❌ Lỗi tính toán: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
        
    print("🎉 HOÀN THÀNH TẤT CẢ CÔNG VIỆC THÀNH CÔNG!")

if __name__ == "__main__":
    main()
