# -*- coding: utf-8 -*-
"""
Script: calculate_report_gan.py
Optimizes the "Báo cáo gán" report. Reads raw Last Mile data, performs AM/Bưu cục mapping,
writes static summary sheets to a lightweight Google Sheet, captures screenshots with Playwright,
and posts summaries with styled table images to GTalk.
"""

import os
import sys
import time
import json
import requests
import pandas as pd
import numpy as np
from datetime import datetime
import pytz
import gspread
from google.oauth2.service_account import Credentials
from playwright.sync_api import sync_playwright
from PIL import Image
import urllib3

import socket
socket.setdefaulttimeout(30)

# Suppress SSL warnings due to verify=False on internal API calls
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
SOURCE_SPREADSHEET_ID = "1DuMW_ajrtrmLlMNslJY2UIMWygVY1cFD4QhKnX9YGNQ"
BACKUP_COCAUMOI_SPREADSHEET_ID = "1x8MxOZV0wMFi7NmXlMaxjBbWjr6zyylUE2rjI4votmw"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(BASE_DIR, "config_gan.json")
SNAPSHOT_FILE = os.path.join(BASE_DIR, "snapshot_gan.json")

SERVICE_ACCOUNT_CANDIDATES = [
    r"C:\Users\lap4all\Desktop\Backlog_Automation\credentials.json",
    r"C:\Users\lap4all\Downloads\credentials.json",
    r"C:\Users\lap4all\Downloads\service_account.json",
    r"C:\Users\lap4all\Desktop\credentials.json",
    os.path.join(BASE_DIR, "credentials.json"),
    os.path.join(BASE_DIR, "service_account.json"),
]

# Load dotenv if exists
from dotenv import load_dotenv
env_path = r"c:\Users\lap4all\Desktop\New folder\.env"
if os.path.exists(env_path):
    load_dotenv(dotenv_path=env_path, override=True)
else:
    load_dotenv()

GTALK_OA_TOKEN = os.environ.get("GAN_GTALK_OA_TOKEN") or "2067164759497973760:RfgqBJY4QtV18udu1wMEfhmoRNI4hgBv"
GTALK_CHANNEL_ID = os.environ.get("GAN_GTALK_CHANNEL_ID") or "2073929320358825984"

def find_service_account_file():
    for p in SERVICE_ACCOUNT_CANDIDATES:
        if os.path.isfile(p):
            return p
    raise FileNotFoundError("Không tìm thấy file credentials.json.")

# ============ GOOGLE SHEETS SETUP & CACHING ============
def get_gspread_client(spreadsheet_id=None):
    scopes = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]
    # Prioritize authorized_user credentials first for fast auth without PermissionErrors
    auth_user_candidates = [
        os.path.join(BASE_DIR, 'authorized_user.json'),
        r'C:\Users\lap4all\Documents\Auto report\authorized_user.json',
        r'C:\Users\lap4all\Desktop\Backlog_Automation\authorized_user.json',
        'authorized_user.json'
    ]
    for auth_user_file in auth_user_candidates:
        if os.path.exists(auth_user_file):
            try:
                from google.oauth2.credentials import Credentials as UserCredentials
                creds = UserCredentials.from_authorized_user_file(auth_user_file, scopes=scopes)
                gc = gspread.authorize(creds)
                if spreadsheet_id:
                    gc.open_by_key(spreadsheet_id)
                print(f"✔️ Đã xác thực thành công qua {auth_user_file}")
                return gc
            except Exception as e:
                err_msg = str(e) or repr(e)
                print(f"⚠️ Authorized user ({auth_user_file}) thất bại: {err_msg}")

    for cred_path in SERVICE_ACCOUNT_CANDIDATES:
        if os.path.isfile(cred_path):
            try:
                creds = Credentials.from_service_account_file(cred_path, scopes=scopes)
                gc = gspread.authorize(creds)
                if spreadsheet_id:
                    gc.open_by_key(spreadsheet_id)
                return gc
            except Exception as e:
                err_msg = str(e) or repr(e)
                print(f"⚠️ Service account ({cred_path}) không thể mở sheet: {err_msg}")

    raise PermissionError("Không thể xác thực Google Sheets bằng credentials.json hoặc authorized_user.json")

def get_credentials():
    return get_gspread_client()


def get_or_create_target_sheet(gc):
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r") as f:
                config = json.load(f)
                sheet_id = config.get("target_sheet_id")
                if sheet_id:
                    sh = gc.open_by_key(sheet_id)
                    print(f"✔️ Đã mở Google Sheet báo cáo gán đã tồn tại: https://docs.google.com/spreadsheets/d/{sheet_id}")
                    return sh
        except Exception as e:
            print(f"⚠️ Không mở được sheet theo config cũ, tiến hành tạo mới. Chi tiết: {e}")

    print("🆕 Đang tạo Google Sheet báo cáo gán mới...")
    sh = gc.create("Báo cáo gán tối ưu")
    sheet_id = sh.id
    
    # Share spreadsheet publicly (anyone with the link can view)
    try:
        sh.share(None, perm_type='anyone', role='reader')
        print("🔓 Đã chia sẻ công khai Google Sheet (ai có link cũng có thể xem).")
    except Exception as e:
        print(f"⚠️ Lỗi chia sẻ công khai: {e}")

    with open(CONFIG_FILE, "w") as f:
        json.dump({"target_sheet_id": sheet_id}, f)
        
    print(f"✔️ Đã tạo thành công Google Sheet mới: https://docs.google.com/spreadsheets/d/{sheet_id}")
    return sh

# ============ FORMATTING HELPERS ============
def make_color(hex_str):
    hex_str = hex_str.lstrip('#')
    r = int(hex_str[0:2], 16) / 255.0
    g = int(hex_str[2:4], 16) / 255.0
    b = int(hex_str[4:6], 16) / 255.0
    return {"red": r, "green": g, "blue": b}

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

# ============ DATA CALCULATION ============
def calculate_summary(df):
    """Tính toán dữ liệu tổng hợp theo Bưu cục và AM"""
    df = df.copy()
    # 1. Bưu cục Level
    # layVolume, layChuaGan, layDaGan
    df['is_lay'] = df['Loại đơn'] == 'Lấy'
    df['is_giao'] = (df['Loại đơn'] != 'Lấy') & (df['Loại đơn'] != 'Trả')
    df['is_chua_gan'] = df['Trạng thái'] == 'Chưa có chuyến đi trong ngày'
    df['is_da_gan'] = df['Trạng thái'].isin(['Đang có chuyến đi trong ngày', 'Đã có chuyến đi trong ngày'])

    bc_grp = df.groupby(['AM', 'Bưu cục']).agg(
        layVolume=('is_lay', 'sum'),
        layChuaGan=('is_lay', lambda x: (x & df.loc[x.index, 'is_chua_gan']).sum()),
        layDaGan=('is_lay', lambda x: (x & df.loc[x.index, 'is_da_gan']).sum()),
        giaoVolume=('is_giao', 'sum'),
        giaoChuaGan=('is_giao', lambda x: (x & df.loc[x.index, 'is_chua_gan']).sum()),
        giaoDaGan=('is_giao', lambda x: (x & df.loc[x.index, 'is_da_gan']).sum())
    ).reset_index()

    bc_grp['layTiLe'] = np.where(bc_grp['layVolume'] > 0, bc_grp['layChuaGan'] / bc_grp['layVolume'], 0.0)
    bc_grp['giaoTiLe'] = np.where(bc_grp['giaoVolume'] > 0, bc_grp['giaoChuaGan'] / bc_grp['giaoVolume'], 0.0)

    # Sort Bưu cục table: AM ASC, Bưu cục ASC
    bc_grp = bc_grp.sort_values(by=['AM', 'Bưu cục']).reset_index(drop=True)

    # 2. AM Level
    am_grp = df.groupby('AM').agg(
        layVolume=('is_lay', 'sum'),
        layChuaGan=('is_lay', lambda x: (x & df.loc[x.index, 'is_chua_gan']).sum()),
        layDaGan=('is_lay', lambda x: (x & df.loc[x.index, 'is_da_gan']).sum()),
        giaoVolume=('is_giao', 'sum'),
        giaoChuaGan=('is_giao', lambda x: (x & df.loc[x.index, 'is_chua_gan']).sum()),
        giaoDaGan=('is_giao', lambda x: (x & df.loc[x.index, 'is_da_gan']).sum())
    ).reset_index()

    am_grp['layTiLe'] = np.where(am_grp['layVolume'] > 0, am_grp['layChuaGan'] / am_grp['layVolume'], 0.0)
    am_grp['giaoTiLe'] = np.where(am_grp['giaoVolume'] > 0, am_grp['giaoChuaGan'] / am_grp['giaoVolume'], 0.0)

    # Sort AM table: assignment rate ASC (tỷ lệ gán thấp nhất lên đầu)
    am_grp['tongVolume'] = am_grp['layVolume'] + am_grp['giaoVolume']
    am_grp['tongDaGan'] = am_grp['layDaGan'] + am_grp['giaoDaGan']
    am_grp['tiLeGan'] = np.where(am_grp['tongVolume'] > 0, am_grp['tongDaGan'] / am_grp['tongVolume'], 1.0)
    am_grp = am_grp.sort_values(by=['tiLeGan', 'tongVolume'], ascending=[True, False]).drop(columns=['tongVolume', 'tongDaGan', 'tiLeGan']).reset_index(drop=True)

    # 3. Totals
    totals = {
        'layVolume': int(df['is_lay'].sum()),
        'layChuaGan': int((df['is_lay'] & df['is_chua_gan']).sum()),
        'layDaGan': int((df['is_lay'] & df['is_da_gan']).sum()),
        'giaoVolume': int(df['is_giao'].sum()),
        'giaoChuaGan': int((df['is_giao'] & df['is_chua_gan']).sum()),
        'giaoDaGan': int((df['is_giao'] & df['is_da_gan']).sum())
    }
    totals['layTiLe'] = totals['layChuaGan'] / totals['layVolume'] if totals['layVolume'] > 0 else 0.0
    totals['giaoTiLe'] = totals['giaoChuaGan'] / totals['giaoVolume'] if totals['giaoVolume'] > 0 else 0.0

    return bc_grp, am_grp, totals

# ============ SNAPSHOT MANAGER ============
def get_current_time_slot():
    tz = pytz.timezone("Asia/Ho_Chi_Minh")
    now = datetime.now(tz)
    h = now.hour
    if h < 11:
        return "10h"
    elif h < 13:
        return "12h"
    else:
        return "14h"

def manage_snapshot(am_grp, totals, slot):
    import unicodedata
    snapshot_data = {}
    tz = pytz.timezone("Asia/Ho_Chi_Minh")
    today_str = datetime.now(tz).strftime('%d/%m/%Y')
    
    if os.path.exists(SNAPSHOT_FILE):
        try:
            with open(SNAPSHOT_FILE, "r", encoding="utf-8") as f:
                loaded_data = json.load(f)
                # Chỉ giữ lại các mốc chạy của ngày hôm nay
                for k, v in loaded_data.items():
                    if isinstance(v, dict) and v.get("date") == today_str:
                        # Chuẩn hóa tên AM trong amData sang NFC để tránh lệch Unicode dựng sẵn/tổ hợp
                        normalized_am_data = {}
                        for am_name, am_val in v.get("amData", {}).items():
                            norm_name = unicodedata.normalize('NFC', am_name).strip()
                            normalized_am_data[norm_name] = am_val
                        v["amData"] = normalized_am_data
                        snapshot_data[k] = v
        except Exception:
            pass

    # Save current run (chuẩn hóa tên AM sang NFC)
    current_snap = {
        "date": today_str,
        "totals": {
            "layChuaGan": int(totals["layChuaGan"]),
            "giaoChuaGan": int(totals["giaoChuaGan"])
        },
        "amData": {
            unicodedata.normalize('NFC', row["AM"]).strip(): {
                "layChuaGan": int(row["layChuaGan"]),
                "giaoChuaGan": int(row["giaoChuaGan"])
            }
            for _, row in am_grp.iterrows()
        }
    }
    snapshot_data[slot] = current_snap
    
    with open(SNAPSHOT_FILE, "w", encoding="utf-8") as f:
        json.dump(snapshot_data, f, ensure_ascii=False)

    import re
    valid_slots = []
    for k in snapshot_data.keys():
        if k == slot:
            continue
        if re.match(r"^\d+h(\d+)?$", k.lower()):
            valid_slots.append(k)
            
    if not valid_slots:
        return None, None
        
    def parse_slot(s):
        m = re.match(r"(\d+)h(\d*)", s.lower())
        h = int(m.group(1))
        mins = int(m.group(2)) if m.group(2) else 0
        return h * 60 + mins

    valid_slots.sort(key=parse_slot)
    
    try:
        current_time = parse_slot(slot)
        prev_slots = [s for s in valid_slots if parse_slot(s) < current_time]
        if prev_slots:
            closest_slot = prev_slots[-1]
            return closest_slot, snapshot_data[closest_slot]
    except Exception:
        pass
        
    return None, None

def retry_gspread(max_retries=5, delay=1.5):
    def decorator(func):
        import functools
        import time
        from gspread.exceptions import APIError
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            for attempt in range(max_retries):
                try:
                    return func(*args, **kwargs)
                except APIError as e:
                    code = e.response.status_code if hasattr(e, 'response') and hasattr(e.response, 'status_code') else None
                    print(f"⚠️ Google Sheets API Error (status={code}) during {func.__name__}: {e}")
                    if attempt < max_retries - 1:
                        sleep_time = round(min(5, delay * (1.5 ** attempt)), 1)
                        print(f"🔄 Đang thử lại sau {sleep_time} giây... (Lần thử {attempt + 1}/{max_retries})")
                        time.sleep(sleep_time)
                    else:
                        raise e
                except Exception as e:
                    print(f"⚠️ Lỗi không xác định trong {func.__name__}: {e}")
                    if attempt < max_retries - 1:
                        sleep_time = round(min(5, delay * (1.5 ** attempt)), 1)
                        print(f"🔄 Đang thử lại sau {sleep_time} giây... (Lần thử {attempt + 1}/{max_retries})")
                        time.sleep(sleep_time)
                    else:
                        raise e
        return wrapper
    return decorator

# ============ WRITE & STYLE WORKSHEET ============
@retry_gspread(max_retries=5, delay=3)
def write_sheet_data(sh, ws_name, bc_df, am_df, totals, today_str, time_label, morning_snap=None, prev_slot="10h"):
    # Get/Create Worksheet
    try:
        ws = sh.worksheet(ws_name)
        ws.clear()
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=ws_name, rows="1000", cols="26")

    grid_values = []
    # Pad empty rows/cols to helper list
    for _ in range(5):
        grid_values.append([""] * 26)

    # Row 2 (index 1): BÁO CÁO TỒN CHƯA GÁN
    grid_values[1][0] = "BÁO CÁO TỒN CHƯA GÁN"
    grid_values[1][9] = today_str

    # Row 3 (index 2): Mốc thời gian
    grid_values[2][0] = time_label

    # Row 4 (index 3): Headers level 1
    grid_values[3][0] = "AM"
    grid_values[3][1] = "Bưu Cục"
    grid_values[3][2] = "Hàng Lấy"
    grid_values[3][6] = "Hàng Giao"

    # Row 5 (index 4): Headers level 2
    grid_values[4][0:10] = ["AM", "Bưu Cục", "Volume", "Chưa Gán", "Đã Gán", "Tỉ lệ CG", "Volume", "Chưa Gán", "Đã Gán", "Tỉ lệ CG"]

    # Row 6 onwards (index 5): Left Table Data
    start_row = 5
    bc_count = len(bc_df)
    for idx, row in bc_df.iterrows():
        r_vals = [
            row["AM"], row["Bưu cục"],
            int(row["layVolume"]), int(row["layChuaGan"]), int(row["layDaGan"]), f"{float(row['layTiLe'])*100:.2f}%",
            int(row["giaoVolume"]), int(row["giaoChuaGan"]), int(row["giaoDaGan"]), f"{float(row['giaoTiLe'])*100:.2f}%"
        ]
        grid_values.append(r_vals + [""] * 16)

    # Left Table Total row
    total_row_idx = start_row + bc_count
    left_totals = [
        "TỔNG CỘNG", "",
        totals["layVolume"], totals["layChuaGan"], totals["layDaGan"], f"{totals['layTiLe']*100:.2f}%",
        totals["giaoVolume"], totals["giaoChuaGan"], totals["giaoDaGan"], f"{totals['giaoTiLe']*100:.2f}%"
    ]
    grid_values.append(left_totals + [""] * 16)

    # ===== BẢNG PHẢI: AM Summary (L-T, index 11-19) =====
    grid_values[1][11] = "TỔNG HỢP THEO AM"
    grid_values[1][18] = today_str
    grid_values[2][11] = time_label

    grid_values[3][11] = "AM"
    grid_values[3][12] = "Hàng Lấy"
    grid_values[3][16] = "Hàng Giao"
    grid_values[4][11:20] = ["AM", "Volume", "Chưa Gán", "Đã Gán", "Tỉ lệ CG", "Volume", "Chưa Gán", "Đã Gán", "Tỉ lệ CG"]

    # Populate AM Data
    am_count = len(am_df)
    for i in range(am_count):
        row = am_df.iloc[i]
        r_vals = [
            row["AM"],
            int(row["layVolume"]), int(row["layChuaGan"]), int(row["layDaGan"]), f"{float(row['layTiLe'])*100:.2f}%",
            int(row["giaoVolume"]), int(row["giaoChuaGan"]), int(row["giaoDaGan"]), f"{float(row['giaoTiLe'])*100:.2f}%"
        ]
        for col_offset, val in enumerate(r_vals):
            grid_values[start_row + i][11 + col_offset] = val

    # Populate AM Total row
    for col_offset, val in enumerate(["TỔNG CỘNG", totals["layVolume"], totals["layChuaGan"], totals["layDaGan"], f"{totals['layTiLe']*100:.2f}%", totals["giaoVolume"], totals["giaoChuaGan"], totals["giaoDaGan"], f"{totals['giaoTiLe']*100:.2f}%"]):
        grid_values[start_row + am_count][11 + col_offset] = val

    # AM delta comparison (U-V, index 20-21) if morning snapshot is available
    if morning_snap:
        grid_values[3][20] = f"Δ CG Lấy (vs {prev_slot})"
        grid_values[3][21] = f"Δ CG Giao (vs {prev_slot})"
        
        snap_am = morning_snap.get("amData", {})
        for i in range(am_count):
            am_name = am_df.loc[i, "AM"]
            old = snap_am.get(am_name, {"layChuaGan": 0, "giaoChuaGan": 0})
            delta_lay = int(am_df.loc[i, "layChuaGan"] - old.get("layChuaGan", 0))
            delta_giao = int(am_df.loc[i, "giaoChuaGan"] - old.get("giaoChuaGan", 0))
            grid_values[start_row + i][20] = f"+{delta_lay}" if delta_lay > 0 else str(delta_lay)
            grid_values[start_row + i][21] = f"+{delta_giao}" if delta_giao > 0 else str(delta_giao)

        # Delta Grand Total
        delta_lay_tot = int(totals["layChuaGan"] - morning_snap.get("totals", {}).get("layChuaGan", 0))
        delta_giao_tot = int(totals["giaoChuaGan"] - morning_snap.get("totals", {}).get("giaoChuaGan", 0))
        grid_values[start_row + am_count][20] = f"+{delta_lay_tot}" if delta_lay_tot > 0 else str(delta_lay_tot)
        grid_values[start_row + am_count][21] = f"+{delta_giao_tot}" if delta_giao_tot > 0 else str(delta_giao_tot)

    # Write data directly to Google Sheet in 1 sub-second API call
    print(f"-> Đang ghi bảng dữ liệu mốc '{time_label}' vào tab '{ws_name}'...")
    max_len = max(len(r) for r in grid_values)
    clean_grid = []
    for r in grid_values:
        if len(r) < max_len:
            r = r + [""] * (max_len - len(r))
        clean_grid.append(r)

    end_col_letter = gspread.utils.rowcol_to_a1(1, len(clean_grid[0])).split("1")[0]
    ws.update(range_name=f"A1:{end_col_letter}{len(clean_grid)}", values=clean_grid, value_input_option="USER_ENTERED")

    last_row_cap = start_row + am_count + 1
    last_col_cap = "V" if morning_snap else "T"
    return f"L1:{last_col_cap}{last_row_cap}"

@retry_gspread(max_retries=5, delay=3)
def write_detail_sheet_data(sh, ws_name, df, is_unassigned=True):
    """Tối ưu tốc độ ghi 50,000+ dòng chi tiết lên Google Sheets:
    1. Chuyển đổi DataFrame qua values.tolist()
    2. Dùng RAW input_option tránh parse formula từng ô
    3. Chỉ format Header + Column Widths, bỏ bớt border/conditional format nặng trên 50k dòng
    """
    try:
        ws = sh.worksheet(ws_name)
        ws.clear()
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=ws_name, rows="1000", cols="8")

    headers = ['AM', 'Bưu Cục', 'Mã Đơn Hàng', 'Loại Đơn', 'Khách Hàng', 'Phường/Xã', 'Trạng Thái', 'Thời Gian Tồn Đọng']
    if df.empty:
        grid_values = [headers, ["Không có đơn hàng nào"] + [""] * 7]
    else:
        detail_df = df[['AM', 'Bưu cục', 'Mã đơn hàng', 'Loại đơn', 'Khách hàng', 'Phường/xã', 'Trạng thái', 'Thời gian tồn đọng']].copy()
        detail_df = detail_df.fillna("")
        detail_df = detail_df.sort_values(by=['AM', 'Bưu cục', 'Khách hàng']).reset_index(drop=True)
        grid_values = [headers] + detail_df.astype(str).values.tolist()

    total_rows = len(grid_values)
    print(f"-> Đang ghi siêu tốc {total_rows - 1:,} dòng chi tiết vào tab '{ws_name}'...")
    try:
        ws.resize(rows=max(100, total_rows + 10), cols=8)
    except Exception:
        pass

    # Ghi toàn bộ 25,000+ dòng chi tiết trong 1 API Call duy nhất (RAW mode)
    ws.update(range_name=f"A1:H{total_rows}", values=grid_values, value_input_option="RAW")

    header_color = "#1F618D" if is_unassigned else "#566573"
    requests = [
        cell_format_request(ws.id, 0, 1, 0, 8, {
            "backgroundColor": make_color(header_color),
            "textFormat": {"fontSize": 10, "bold": True, "foregroundColor": make_color("#FFFFFF"), "fontFamily": "Arial"},
            "horizontalAlignment": "CENTER",
            "verticalAlignment": "MIDDLE"
        }),
        row_height_request(ws.id, 0, 1, 28)
    ]
    widths = {0: 160, 1: 220, 2: 120, 3: 80, 4: 100, 5: 160, 6: 220, 7: 120}
    for c_idx, w in widths.items():
        requests.append(col_width_request(ws.id, c_idx, c_idx + 1, w))

    try:
        sh.batch_update({"requests": requests})
    except Exception as e:
        print(f"⚠️ Lỗi định dạng tab chi tiết: {e}")

    return f"https://docs.google.com/spreadsheets/d/{sh.id}/edit#gid={ws.id}"

def send_photo_gtalk(image_path, caption):
    if not GTALK_OA_TOKEN or not GTALK_CHANNEL_ID:
        print("⚠️ Không có GTALK configuration. Bỏ qua gửi GTalk.")
        return False

    print(f"📡 Đang tải ảnh '{os.path.basename(image_path)}' gửi tới GTalk group...")
    try:
        img = Image.open(image_path)
        width, height = img.size
        file_size = os.path.getsize(image_path)
        with open(image_path, "rb") as f:
            file_bytes = f.read()
    except Exception as e:
        print(f"❌ Lỗi đọc file ảnh: {e}")
        return False

    # 1. Initiate Upload
    init_url = "https://mbff.ghn.vn/api/gtalk/initiate-upload"
    payload = {
        "ChannelId": GTALK_CHANNEL_ID,
        "FileName": os.path.basename(image_path),
        "FileSize": str(file_size),
        "MimeType": "image/png",
        "Metadata": json.dumps({"width": width, "height": height}),
        "oaToken": GTALK_OA_TOKEN
    }
    
    try:
        res = requests.post(init_url, json=payload, timeout=20, verify=False)
        if res.status_code != 200 or res.json().get("errorCode") != "success":
            print(f"❌ Initiate upload thất bại: {res.text}")
            return False
        
        presigned_url = res.json()["data"]["PresignedURL"]
        upload_id = res.json()["data"]["UploadId"]
        
        # 2. PUT to S3
        res_put = requests.put(presigned_url, data=file_bytes, headers={"Content-Type": "image/png"}, timeout=60, verify=False)
        if res_put.status_code != 200:
            print("❌ PUT file to S3 thất bại.")
            return False
            
        # 3. Complete Upload
        comp_url = "https://mbff.ghn.vn/api/gtalk/complete-upload"
        res_comp = requests.post(comp_url, json={"oaToken": GTALK_OA_TOKEN, "UploadId": upload_id}, timeout=20, verify=False)
        if res_comp.status_code != 200 or res_comp.json().get("errorCode") != "success":
            print("❌ Complete upload thất bại.")
            return False
            
        file_id = res_comp.json()["data"]["Id"]
        
        # 4. Send message
        send_url = "https://mbff.ghn.vn/api/gtalk/send-message"
        client_msg_id = str(int(time.time() * 1000))
        send_payload = {
            "channelId": GTALK_CHANNEL_ID,
            "clientMsgId": client_msg_id,
            "content": {
                "parseMode": "HTML",
                "attachment": {
                    "caption": caption,
                    "items": [
                        {
                            "image": {
                                "fileId": file_id,
                                "width": width,
                                "height": height
                            }
                        }
                    ]
                }
            },
            "oaToken": GTALK_OA_TOKEN
        }
        
        res_send = requests.post(send_url, json=send_payload, timeout=20, verify=False)
        if res_send.status_code == 200 and res_send.json().get("errorCode") == "success":
            print("✅ Đã gửi ảnh thành công qua GTalk!")
            return True
        else:
            print(f"❌ Gửi tin nhắn GTalk thất bại: {res_send.text}")
    except Exception as e:
        print(f"❌ Kết nối GTalk API lỗi: {e}")
    return False



# ============ MAIN PIPELINE ============
@retry_gspread(max_retries=5, delay=3)
def load_cocaumoi(gc):
    sh_src = gc.open_by_key(SOURCE_SPREADSHEET_ID)
    try:
        ws_co = sh_src.worksheet("cocaumoi")
        co_data = ws_co.get_all_values()
        print("✔️ Đã đọc bảng 'cocaumoi' từ spreadsheet mục tiêu mới.")
    except gspread.exceptions.WorksheetNotFound:
        print("⚠️ Không tìm thấy tab 'cocaumoi' ở sheet mới, đang đọc từ backup sheet gốc...")
        sh_backup = gc.open_by_key(BACKUP_COCAUMOI_SPREADSHEET_ID)
        ws_co = sh_backup.worksheet("cocaumoi")
        co_data = ws_co.get_all_values()
        print("✔️ Đã đọc bảng 'cocaumoi' thành công từ backup sheet gốc.")
    return co_data, sh_src

def main():
    print("================== BẮT ĐẦU PIPELINE BÁO CÁO GÁN TỐI ƯU ==================")
    import unicodedata
    temp_ghn_path = r"c:\Users\lap4all\.gemini\antigravity-ide\scratch\temp_ghn.xlsx"
    
    print("🔑 Đang xác thực thông tin tài khoản Google...")
    try:
        gc = get_gspread_client(SOURCE_SPREADSHEET_ID)
        sh_src = gc.open_by_key(SOURCE_SPREADSHEET_ID)
    except Exception as e:
        print(f"❌ Lỗi xác thực credentials: {e}")
        sys.exit(1)



    if "--no-download" in sys.argv:
        print("⚠️ Bỏ qua bước tải dữ liệu từ GHN. Sử dụng dữ liệu temp_ghn.xlsx hiện có.")
    else:
        print("📥 Đang tự động kích hoạt tải dữ liệu Last Mile mới nhất từ GHN...")
        import subprocess
        download_script = r"c:\Users\lap4all\.gemini\antigravity-ide\scratch\download_report_thuy.py"
        try:
            python_exe = sys.executable or "python"
            result = subprocess.run([python_exe, download_script, "--force"], capture_output=True, text=True, encoding='utf-8', check=True)
            print("✔️ Tải dữ liệu thành công!")
            if result.stdout:
                for line in result.stdout.split('\n'):
                    if "đơn hàng" in line or "HOÀN THÀNH" in line:
                        print(f"  [GHN Bot] {line}")
        except subprocess.CalledProcessError as e:
            print(f"❌ Lỗi tải dữ liệu Last Mile: {e}")
            if not os.path.exists(temp_ghn_path):
                print("❌ Không có file temp_ghn.xlsx làm dự phòng. Dừng script.")
                sys.exit(1)
            print("⚠️ Sử dụng file dữ liệu temp_ghn.xlsx cũ sẵn có trên máy làm dự phòng.")

    # 1. Read cocaumoi
    print("📖 Đang đọc bảng ánh xạ 'cocaumoi'...")
    try:
        co_data, sh_src = load_cocaumoi(gc)
        co_df = pd.DataFrame(co_data[1:], columns=co_data[0])
        co_df['warehouse_id'] = co_df['warehouse_id'].astype(str).str.strip()
        if 'AM' in co_df.columns:
            co_df['AM'] = co_df['AM'].astype(str).str.strip().apply(lambda x: unicodedata.normalize('NFC', x))
        if 'Bưu cục' in co_df.columns:
            co_df['Bưu cục'] = co_df['Bưu cục'].astype(str).str.strip().apply(lambda x: unicodedata.normalize('NFC', x))
    except Exception as e:
        print(f"❌ Lỗi đọc mapping Cơ cấu: {e}")
        sys.exit(1)

    # 2. Read raw excel
    print("📖 Đang tải dữ liệu thô từ temp_ghn.xlsx...")
    try:
        required_cols = ['Mã bưu cục', 'Mã đơn hàng', 'Loại đơn', 'Khách hàng', 'Phường/xã', 'Trạng thái', 'Thời gian tồn đọng']
        raw_df = pd.read_excel(temp_ghn_path, skiprows=1, usecols=required_cols).fillna("")
        raw_df['Mã bưu cục'] = raw_df['Mã bưu cục'].astype(str).str.strip()
    except Exception as e:
        print(f"❌ Lỗi đọc file raw excel: {e}")
        sys.exit(1)

    # Merge raw data with cocaumoi to get AM column
    df_merged = pd.merge(raw_df, co_df, left_on='Mã bưu cục', right_on='warehouse_id', how='left')
    df_merged['AM'] = df_merged['AM'].fillna("Không xác định").astype(str).str.strip().apply(lambda x: unicodedata.normalize('NFC', x))
    df_merged['Bưu cục'] = df_merged['Bưu cục'].fillna("Bưu cục " + df_merged['Mã bưu cục']).astype(str).str.strip().apply(lambda x: unicodedata.normalize('NFC', x))

    # 3. Đẩy toàn bộ dữ liệu thô (Cols A:G) và AM mapping (Col H) sang tab 'LM' trên Google Sheets
    print("📤 Đang ghi đè dữ liệu thô & AM mapping vào tab 'LM' trên Google Sheets...")
    try:
        try:
            ws_lm = sh_src.worksheet("LM")
        except gspread.exceptions.WorksheetNotFound:
            ws_lm = sh_src.add_worksheet(title="LM", rows="1000", cols="8")

        headers_lm = ['Mã bưu cục', 'Mã đơn hàng', 'Loại đơn', 'Khách hàng', 'Phường/xã', 'Trạng thái', 'Thời gian tồn đọng', 'AM']
        grid_lm = [headers_lm] + df_merged[headers_lm].astype(str).values.tolist()

        try:
            ws_lm.resize(rows=len(grid_lm) + 5, cols=8)
        except Exception:
            pass

        ws_lm.update(range_name=f"A1:H{len(grid_lm)}", values=grid_lm, value_input_option="RAW")
        print(f"✔️ Đã đẩy thành công {len(df_merged):,} dòng sang tab 'LM' (Cols A:G dữ liệu thô + Col H AM mapping)!")
    except Exception as e:
        print(f"⚠️ Không thể cập nhật tab 'LM': {e}")

    # Determine time slot & date label
    tz = pytz.timezone("Asia/Ho_Chi_Minh")
    now = datetime.now(tz)
    today_str = now.strftime('%d/%m/%Y')
    # Filter out flags starting with '--'
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    if args:
        slot = args[0]
        time_label = f"Mốc {slot}"
    else:
        # Lấy mốc giờ phút thực tế làm slot (ví dụ: 17h31) thay vì cố định 10h/12h/14h
        slot = now.strftime('%Hh%M').lower()
        if slot.startswith('0'):
            slot = slot[1:]
        time_label = f"Mốc {slot}"
    
    print(f"⏰ Chạy mốc: {time_label} ngày {today_str}")

    # 3. Calculate statistics
    print("📊 Đang phân tích và tính toán dữ liệu...")
    # Báo cáo 1: Chung (Full)
    bc_full, am_full, tot_full = calculate_summary(df_merged)

    # Báo cáo 2: TTS (customer == 'TTS')
    df_tts = df_merged[df_merged['Khách hàng'] == 'TTS']
    bc_tts, am_tts, tot_tts = calculate_summary(df_tts)

    # Báo cáo 3: SPE (customer == 'SHOPEE')
    df_spe = df_merged[df_merged['Khách hàng'] == 'SHOPEE']
    bc_spe, am_spe, tot_spe = calculate_summary(df_spe)

    # 4. Manage snapshot & load previous snapshot
    prev_slot, morning_snap = manage_snapshot(am_full, tot_full, slot)

    # 5. Connect and update target sheet (use existing spreadsheet directly to avoid Drive storage quota error)
    sh_target = sh_src

    print("📤 Đang cập nhật dữ liệu báo cáo lên Google Sheets...")
    range_full = write_sheet_data(sh_target, "Báo Cáo", bc_full, am_full, tot_full, today_str, time_label, morning_snap, prev_slot)
    range_tts = write_sheet_data(sh_target, "[TTs]Báo Cáo", bc_tts, am_tts, tot_tts, today_str, time_label, morning_snap, prev_slot)
    range_spe = write_sheet_data(sh_target, "[SPX]Báo Cáo", bc_spe, am_spe, tot_spe, today_str, time_label, morning_snap, prev_slot)
    
    print("📤 Đang ghi đè danh sách đơn chi tiết Chưa Gán lên Google Sheets...")
    df_chua_gan = df_merged[df_merged['Trạng thái'] == 'Chưa có chuyến đi trong ngày']
    
    url_chua_gan = write_detail_sheet_data(sh_target, "[Chi Tiết] Chưa Gán", df_chua_gan, is_unassigned=True)
    
    # Bỏ tab [Chi Tiết] Đã Gán nếu tồn tại trên Google Sheets để tối ưu bộ nhớ
    try:
        ws_da_gan = sh_target.worksheet("[Chi Tiết] Đã Gán")
        sh_target.del_worksheet(ws_da_gan)
        print("🗑️ Đã xóa tab '[Chi Tiết] Đã Gán' dư thừa trên Google Sheets.")
    except Exception:
        pass
    
    print("✅ Đã cập nhật xong tất cả các sheet báo cáo tổng hợp và chi tiết!")

    # 6. Playwright screenshot and broadcast
    # Build HTML table for Playwright element screenshot
    def make_table_html(title, am_df, totals, is_delta=False, slot_snap=None, delta_label="10h"):
        body_rows = ""
        for idx, row in am_df.iterrows():
            lay_chua_gan = int(row['layChuaGan'])
            lay_vol = int(row['layVolume'])
            giao_chua_gan = int(row['giaoChuaGan'])
            giao_vol = int(row['giaoVolume'])
            
            lay_rate = f"{lay_chua_gan / lay_vol * 100:.2f}%" if lay_vol > 0 else "0.00%"
            giao_rate = f"{giao_chua_gan / giao_vol * 100:.2f}%" if giao_vol > 0 else "0.00%"
            
            delta_html = ""
            if is_delta and slot_snap:
                snap_am = slot_snap.get("amData", {})
                old = snap_am.get(row['AM'], {"layChuaGan": 0, "giaoChuaGan": 0})
                d_lay = lay_chua_gan - old.get("layChuaGan", 0)
                d_giao = giao_chua_gan - old.get("giaoChuaGan", 0)
                
                def fmt_d(val):
                    if val > 0:
                        return f"<td class='delta-up'>▲ +{val:,}</td>"
                    elif val < 0:
                        return f"<td class='delta-down'>▼ {abs(val):,}</td>"
                    else:
                        return "<td>-</td>"
                delta_html = fmt_d(d_lay) + fmt_d(d_giao)
                
            body_rows += f"""
            <tr>
                <td style="text-align: left; font-weight: bold;">{row['AM']}</td>
                <td>{lay_vol:,}</td>
                <td style="color: #c2410c; font-weight: bold;">{lay_chua_gan:,}</td>
                <td>{int(row['layDaGan']):,}</td>
                <td>{lay_rate}</td>
                <td>{giao_vol:,}</td>
                <td style="color: #c2410c; font-weight: bold;">{giao_chua_gan:,}</td>
                <td>{int(row['giaoDaGan']):,}</td>
                <td>{giao_rate}</td>
                {delta_html}
            </tr>
            """
        
        lay_tot_rate = f"{totals['layChuaGan'] / totals['layVolume'] * 100:.2f}%" if totals['layVolume'] > 0 else "0.00%"
        giao_tot_rate = f"{totals['giaoChuaGan'] / totals['giaoVolume'] * 100:.2f}%" if totals['giaoVolume'] > 0 else "0.00%"
        
        delta_total_html = ""
        if is_delta and slot_snap:
            tot_snap = slot_snap.get("totals", {})
            d_lay_tot = totals['layChuaGan'] - tot_snap.get("layChuaGan", 0)
            d_giao_tot = totals['giaoChuaGan'] - tot_snap.get("giaoChuaGan", 0)
            
            def fmt_td(val):
                if val > 0:
                    return f"<td class='delta-up'>▲ +{val:,}</td>"
                elif val < 0:
                    return f"<td class='delta-down'>▼ {abs(val):,}</td>"
                else:
                    return "<td>-</td>"
            delta_total_html = fmt_td(d_lay_tot) + fmt_td(d_giao_tot)
            
        grand_total_row = f"""
        <tr class="grand-total">
            <td style="text-align: left;">TỔNG CỘNG</td>
            <td>{totals['layVolume']:,}</td>
            <td>{totals['layChuaGan']:,}</td>
            <td>{totals['layDaGan']:,}</td>
            <td>{lay_tot_rate}</td>
            <td>{totals['giaoVolume']:,}</td>
            <td>{totals['giaoChuaGan']:,}</td>
            <td>{totals['giaoDaGan']:,}</td>
            <td>{giao_tot_rate}</td>
            {delta_total_html}
        </tr>
        """
        
        delta_header = ""
        if is_delta and slot_snap:
            delta_header = f"""
            <th class="delta-header">Δ CG Lấy<br>(vs {delta_label})</th>
            <th class="delta-header">Δ CG Giao<br>(vs {delta_label})</th>
            """
            
        return f"""
        <table>
            <tr class="title-row">
                <th colspan="{11 if is_delta and slot_snap else 9}">{title} - {today_str} ({time_label})</th>
            </tr>
            <tr class="header-row">
                <th rowspan="2">AM</th>
                <th colspan="4" style="background-color: #2E4053;">HÀNG LẤY</th>
                <th colspan="4" style="background-color: #2E4053;">HÀNG GIAO</th>
                {delta_header}
            </tr>
            <tr class="subheader-row">
                <th>Volume</th>
                <th>Chưa Gán</th>
                <th>Đã Gán</th>
                <th>Tỷ lệ CG</th>
                <th>Volume</th>
                <th>Chưa Gán</th>
                <th>Đã Gán</th>
                <th>Tỷ lệ CG</th>
            </tr>
            {body_rows}
            {grand_total_row}
        </table>
        """

    is_comparison = morning_snap is not None
    html_full = make_table_html("BÁO CÁO CHƯA GÁN CHUNG", am_full, tot_full, is_comparison, morning_snap, prev_slot)
    html_tts = make_table_html("BÁO CÁO CHƯA GÁN TTS", am_tts, tot_tts, is_comparison, morning_snap, prev_slot)
    html_spe = make_table_html("BÁO CÁO CHƯA GÁN SPE (SHOPEE)", am_spe, tot_spe, is_comparison, morning_snap, prev_slot)

    full_html_page = f"""<!DOCTYPE html>
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
      }}
      .table-container {{
        display: inline-block;
        padding: 10px;
        background: #FFFFFF;
        margin-bottom: 25px;
      }}
      table {{
        border-collapse: collapse;
        width: {1000 if is_comparison else 850}px;
        font-size: 13px;
        border: 1px solid #BDC3C7;
      }}
      th, td {{
        border: 1px solid #BDC3C7;
        padding: 8px 10px;
        text-align: center;
        vertical-align: middle;
      }}
      .title-row th {{
        background-color: #1B4F72;
        color: #FFFFFF;
        font-size: 15px;
        font-weight: 700;
        height: 35px;
      }}
      .header-row th, .subheader-row th {{
        background-color: #2C3E50;
        color: #FFFFFF;
        font-weight: bold;
        font-size: 12px;
      }}
      .grand-total {{
        background-color: #D5D8DC;
        font-weight: bold;
      }}
      .delta-up {{
        color: #C0392B;
        font-weight: bold;
        background-color: #FADBD8;
      }}
      .delta-down {{
        color: #27AE60;
        font-weight: bold;
        background-color: #D4EFDF;
      }}
      .delta-header {{
        background-color: #7D3C98 !important;
      }}
    </style>
    </head>
    <body>
      <div id="div-full" class="table-container">{html_full}</div>
      <br/>
      <div id="div-tts" class="table-container">{html_tts}</div>
      <br/>
      <div id="div-spe" class="table-container">{html_spe}</div>
    </body>
    </html>
    """

    html_path = os.path.join(BASE_DIR, "temp_gan_tables.html")
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(full_html_page)

    print("📸 Đang khởi chạy Playwright để chụp hình các bảng...")
    img_paths = {
        "full": os.path.join(BASE_DIR, "table_gan_chung.png"),
        "tts": os.path.join(BASE_DIR, "table_gan_tts.png"),
        "spe": os.path.join(BASE_DIR, "table_gan_spe.png")
    }

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch()
            page = browser.new_page()
            page.set_viewport_size({"width": 1200, "height": 1800})
            page.goto(f"file:///{os.path.abspath(html_path)}")
            page.wait_for_timeout(500)
            
            page.locator("#div-full").screenshot(path=img_paths["full"])
            page.locator("#div-tts").screenshot(path=img_paths["tts"])
            page.locator("#div-spe").screenshot(path=img_paths["spe"])
            
            browser.close()
        print("✔️ Chụp thành công 3 bảng báo cáo.")
    except Exception as e:
        print(f"❌ Lỗi chụp ảnh Playwright: {e}")
        # Cleanup
        if os.path.exists(html_path):
            os.remove(html_path)
        sys.exit(1)

    # Cleanup temp HTML
    if os.path.exists(html_path):
        os.remove(html_path)

    # 7. Broadcast reports to GTalk
    print("📡 Đang phát sóng báo cáo sang GTalk group...")
    sheet_link = f"https://docs.google.com/spreadsheets/d/{sh_target.id}"

    def make_top_am_text(am_df):
        am_calc = am_df.copy()
        am_calc['tongVolume'] = am_calc['layVolume'] + am_calc['giaoVolume']
        am_calc['tongDaGan'] = am_calc['layDaGan'] + am_calc['giaoDaGan']
        am_calc['tongChuaGan'] = am_calc['layChuaGan'] + am_calc['giaoChuaGan']
        am_calc['tiLeGan'] = np.where(am_calc['tongVolume'] > 0, am_calc['tongDaGan'] / am_calc['tongVolume'], 1.0)
        
        am_calc_filtered = am_calc[am_calc['tongVolume'] > 0]
        top_low_am = am_calc_filtered.sort_values(by='tiLeGan', ascending=True).head(5)
        
        text = "\n🏆 <b>TOP 5 AM TỶ LỆ GÁN THẤP NHẤT:</b>\n"
        if top_low_am.empty:
            text += "  (Không có dữ liệu)\n"
        else:
            for r_idx, (_, row) in enumerate(top_low_am.iterrows(), 1):
                ti_le_gan_pct = row['tiLeGan'] * 100
                text += f"  {r_idx}. {row['AM']}: <b>{ti_le_gan_pct:.2f}%</b> (Chưa gán: <span style='color:#c2410c'><b>{int(row['tongChuaGan'])}</b></span>/{int(row['tongVolume'])})\n"
        return text

    top_am_full = make_top_am_text(am_full)
    top_am_tts = make_top_am_text(am_tts)
    top_am_spe = make_top_am_text(am_spe)

    # General Message caption
    caption_full = f"📊 <b>BÁO CÁO CHƯA GÁN CHUNG (TOÀN BỘ)</b>\n" \
                   f"⏱️ <b>Mốc cập nhật:</b> {time_label} ngày {today_str}\n" \
                   f"🔗 <a href=\"{sheet_link}\"><b>chi tiết link</b></a>\n" \
                   f"{top_am_full}\n" \
                   f"📦 <b>HÀNG LẤY:</b>\n" \
                   f"  • Tổng Volume: <b>{tot_full['layVolume']:,}</b>\n" \
                   f"  • Chưa Gán: <span style='color:#c2410c'><b>{tot_full['layChuaGan']:,}</b></span> ({tot_full['layTiLe']*100:.2f}%)\n" \
                   f"  • Đã Gán: <b>{tot_full['layDaGan']:,}</b>\n\n" \
                   f"🚚 <b>HÀNG GIAO:</b>\n" \
                   f"  • Tổng Volume: <b>{tot_full['giaoVolume']:,}</b>\n" \
                   f"  • Chưa Gán: <span style='color:#c2410c'><b>{tot_full['giaoChuaGan']:,}</b></span> ({tot_full['giaoTiLe']*100:.2f}%)\n" \
                   f"  • Đã Gán: <b>{tot_full['giaoDaGan']:,}</b>"

    caption_tts = f"📊 <b>BÁO CÁO CHƯA GÁN KHÁCH HÀNG TTS</b>\n" \
                  f"⏱️ <b>Mốc cập nhật:</b> {time_label} ngày {today_str}\n" \
                  f"🔗 <a href=\"{sheet_link}\"><b>chi tiết link</b></a>\n" \
                  f"{top_am_tts}\n" \
                  f"📦 <b>HÀNG LẤY:</b>\n" \
                  f"  • Tổng Volume: <b>{tot_tts['layVolume']:,}</b>\n" \
                  f"  • Chưa Gán: <span style='color:#c2410c'><b>{tot_tts['layChuaGan']:,}</b></span> ({tot_tts['layTiLe']*100:.2f}%)\n" \
                  f"  • Đã Gán: <b>{tot_tts['layDaGan']:,}</b>\n\n" \
                  f"🚚 <b>HÀNG GIAO:</b>\n" \
                  f"  • Tổng Volume: <b>{tot_tts['giaoVolume']:,}</b>\n" \
                  f"  • Chưa Gán: <span style='color:#c2410c'><b>{tot_tts['giaoChuaGan']:,}</b></span> ({tot_tts['giaoTiLe']*100:.2f}%)\n" \
                  f"  • Đã Gán: <b>{tot_tts['giaoDaGan']:,}</b>"

    caption_spe = f"📊 <b>BÁO CÁO CHƯA GÁN KHÁCH HÀNG SPE (SHOPEE)</b>\n" \
                  f"⏱️ <b>Mốc cập nhật:</b> {time_label} ngày {today_str}\n" \
                  f"🔗 <a href=\"{sheet_link}\"><b>chi tiết link</b></a>\n" \
                  f"{top_am_spe}\n" \
                  f"📦 <b>HÀNG LẤY:</b>\n" \
                  f"  • Tổng Volume: <b>{tot_spe['layVolume']:,}</b>\n" \
                  f"  • Chưa Gán: <span style='color:#c2410c'><b>{tot_spe['layChuaGan']:,}</b></span> ({tot_spe['layTiLe']*100:.2f}%)\n" \
                  f"  • Đã Gán: <b>{tot_spe['layDaGan']:,}</b>\n\n" \
                  f"🚚 <b>HÀNG GIAO:</b>\n" \
                  f"  • Tổng Volume: <b>{tot_spe['giaoVolume']:,}</b>\n" \
                  f"  • Chưa Gán: <span style='color:#c2410c'><b>{tot_spe['giaoChuaGan']:,}</b></span> ({tot_spe['giaoTiLe']*100:.2f}%)\n" \
                  f"  • Đã Gán: <b>{tot_spe['giaoDaGan']:,}</b>"

    # Send 3 report photos and captions to GTalk
    send_photo_gtalk(img_paths["full"], caption_full)
    time.sleep(2)
    send_photo_gtalk(img_paths["tts"], caption_tts)
    time.sleep(2)
    send_photo_gtalk(img_paths["spe"], caption_spe)

    # Cleanup local photos
    for p in img_paths.values():
        if os.path.exists(p):
            try:
                os.remove(p)
            except Exception:
                pass

    print("================== HOÀN THÀNH PIPELINE THÀNH CÔNG VÀ AN TOÀN! ==================")

if __name__ == "__main__":
    main()
