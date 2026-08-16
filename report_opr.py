# -*- coding: utf-8 -*-
"""
Script: report_opr.py
Processes OPR performance data from Google Sheets, updates the REPORT_OPR tab with N vs N-1 comparison,
renders visual tables with modern premium design using Playwright, and broadcasts them to Telegram & GTalk.
"""

import os
import sys
import time
import urllib3
import requests
import unicodedata
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import pytz
from dotenv import load_dotenv
import gspread
from google.oauth2.service_account import Credentials
from playwright.sync_api import sync_playwright
from PIL import Image

# Configure output encoding for Vietnamese characters
try:
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    if hasattr(sys.stderr, 'reconfigure'):
        sys.stderr.reconfigure(encoding='utf-8')
except Exception:
    pass

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Load dotenv from New folder
env_path = r"c:\Users\lap4all\Desktop\New folder\.env"
if os.path.exists(env_path):
    load_dotenv(dotenv_path=env_path, override=True)
else:
    load_dotenv()

# ============ CONFIG & CONSTANTS ============
SPREADSHEET_ID = "1B-QCbEnPpILFFEWPYheGdmkgYV9gSf4lAyQMlhzwOCM"
TELEGRAM_TOKEN = "8570130113:AAGXRiUaKBknVpgtm1_i9ZA47JRjAXmB21M"
TELEGRAM_CHAT_ID = "-5058464865"
GTALK_OA_TOKEN = "2067164759497973760:RfgqBJY4QtV18udu1wMEfhmoRNI4hgBv"
GTALK_CHANNEL_ID = "2067283005274091520"  # Dedicated GTalk Channel ID

SERVICE_ACCOUNT_CANDIDATES = [
    r"C:\Users\lap4all\Desktop\Backlog_Automation\credentials.json",
    r"C:\Users\lap4all\Downloads\credentials.json",
    r"C:\Users\lap4all\Downloads\service_account.json",
    r"C:\Users\lap4all\Desktop\credentials.json",
    "credentials.json",
    "service_account.json",
]

# Override config if defined in env (chỉ dùng OPR_GTALK_OA_TOKEN riêng, không bị ghi đè bởi GTALK_OA_TOKEN dùng chung)
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN") or os.environ.get("TELEGRAM_BOT_TOKEN") or TELEGRAM_TOKEN
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID") or TELEGRAM_CHAT_ID
GTALK_OA_TOKEN = os.environ.get("OPR_GTALK_OA_TOKEN") or GTALK_OA_TOKEN
GTALK_CHANNEL_ID = os.environ.get("OPR_GTALK_CHANNEL_ID") or GTALK_CHANNEL_ID

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def get_gspread_client(spreadsheet_id=SPREADSHEET_ID):
    scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    for cred_path in SERVICE_ACCOUNT_CANDIDATES:
        if os.path.isfile(cred_path):
            try:
                creds = Credentials.from_service_account_file(cred_path, scopes=scopes)
                gc = gspread.authorize(creds)
                if spreadsheet_id:
                    gc.open_by_key(spreadsheet_id)
                return gc
            except Exception as e:
                print(f"⚠️ Service account ({cred_path}) không có quyền: {e}. Đang chuyển sang authorized_user.json...")

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
                return gc
            except Exception as e:
                pass

    raise PermissionError("Không thể xác thực Google Sheets bằng credentials.json hoặc authorized_user.json")


# ============ HELPER FUNCTIONS ============
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

def find_col(headers, candidates, default_idx):
    for idx, h in enumerate(headers):
        h_clean = str(h).strip().lower()
        if h_clean in candidates:
            return idx
    return default_idx

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

def border_request(sheet_id, start_row, end_row, start_col, end_col, color_hex="#D9D9D9"):
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

# ============ TELEGRAM SENDER ============
def send_photo_telegram(image_path, caption=""):
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendPhoto"
    with open(image_path, "rb") as f:
        files = {"photo": f}
        payload = {
            "chat_id": TELEGRAM_CHAT_ID,
            "caption": caption,
            "parse_mode": "HTML",
        }
        resp = requests.post(url, data=payload, files=files)
    result = resp.json()
    if result.get("ok"):
        print("✅ Đã gửi ảnh lên Telegram thành công.")
    else:
        print(f"❌ Lỗi gửi Telegram: {result}")
    return result

# ============ GTALK SENDER ============
def send_photo_gtalk(image_path, caption=""):
    if not GTALK_OA_TOKEN or not GTALK_CHANNEL_ID:
        print("⚠️ Không tìm thấy GTALK_OA_TOKEN hoặc GTALK_CHANNEL_ID. Bỏ qua gửi GTalk.")
        return False

    print("📡 Đang gửi ảnh báo cáo sang GTalk...")
    try:
        img = Image.open(image_path)
        width, height = img.size
        file_size = os.path.getsize(image_path)
    except Exception as e:
        print(f"❌ Lỗi đọc ảnh {image_path}: {e}")
        return False

    # Step 1: Initiate Upload
    initiate_url = "https://mbff.ghn.vn/api/gtalk/initiate-upload"
    payload_init = {
        "ChannelId": GTALK_CHANNEL_ID,
        "FileName": os.path.basename(image_path),
        "FileSize": str(file_size),
        "MimeType": "image/png",
        "Metadata": f'{{"width": {width}, "height": {height}}}',
        "oaToken": GTALK_OA_TOKEN
    }
    headers = {"Content-Type": "application/json"}
    try:
        res_init = requests.post(initiate_url, json=payload_init, headers=headers, timeout=20, verify=False)
        if res_init.status_code != 200:
            print(f"❌ Lỗi initiate upload HTTP {res_init.status_code}: {res_init.text}")
            return False
        res_data = res_init.json()
        if res_data.get("errorCode") != "success":
            print(f"❌ Lỗi initiate upload API: {res_data.get('error')}")
            return False
        
        presigned_url = res_data["data"]["PresignedURL"]
        upload_id = res_data["data"]["UploadId"]
    except Exception as e:
        print(f"❌ Lỗi kết nối khi initiate upload GTalk: {e}")
        return False

    # Step 2: Upload to S3
    try:
        with open(image_path, "rb") as f:
            headers_put = {"Content-Type": "image/png"}
            res_put = requests.put(presigned_url, data=f, headers=headers_put, timeout=60, verify=False)
            if res_put.status_code != 200:
                print(f"❌ Lỗi PUT lên S3 HTTP {res_put.status_code}: {res_put.text}")
                return False
    except Exception as e:
        print(f"❌ Lỗi upload file lên S3 GTalk: {e}")
        return False

    # Step 3: Complete Upload
    complete_url = "https://mbff.ghn.vn/api/gtalk/complete-upload"
    payload_complete = {
        "oaToken": GTALK_OA_TOKEN,
        "UploadId": upload_id
    }
    try:
        res_comp = requests.post(complete_url, json=payload_complete, headers=headers, timeout=20, verify=False)
        if res_comp.status_code != 200:
            print(f"❌ Lỗi complete upload HTTP {res_comp.status_code}: {res_comp.text}")
            return False
        res_data_comp = res_comp.json()
        if res_data_comp.get("errorCode") != "success":
            print(f"❌ Lỗi complete upload API: {res_data_comp.get('error')}")
            return False
        file_id = res_data_comp["data"]["Id"]
    except Exception as e:
        print(f"❌ Lỗi kết nối khi complete upload GTalk: {e}")
        return False

    # Step 4: Send Message
    send_url = "https://mbff.ghn.vn/api/gtalk/send-message"
    client_msg_id = str(int(time.time() * 1000))
    payload_send = {
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
    try:
        res_send = requests.post(send_url, json=payload_send, headers=headers, timeout=20, verify=False)
        if res_send.status_code == 200:
            res_data_send = res_send.json()
            if res_data_send.get("errorCode") == "success":
                print("✅ Đã gửi ảnh báo cáo sang GTalk thành công!")
                return True
            else:
                print(f"❌ Lỗi gửi tin nhắn GTalk API: {res_data_send.get('error')}")
        else:
            print(f"❌ Lỗi HTTP {res_send.status_code}: {res_send.text}")
    except Exception as e:
        print(f"❌ Lỗi kết nối khi gửi tin nhắn GTalk: {e}")
    return False

# ============ MAIN PIPELINE ============
def main():
    print("🔐 Kết nối Google Sheets...")
    client = get_gspread_client(SPREADSHEET_ID)

    print(f"Đang kết nối tới spreadsheet ID: {SPREADSHEET_ID}...")
    try:
        sh = client.open_by_key(SPREADSHEET_ID)
        print(f"✔️ Đã kết nối thành công tới: '{sh.title}'")
    except Exception as e:
        print(f"❌ Lỗi mở spreadsheet: {e}")
        sys.exit(1)

    # 1. Load Data Sheet "OPR"
    try:
        data_sheet = sh.worksheet("OPR")
        raw_values = data_sheet.get_all_values()
        print(f"✔️ Đã đọc dữ liệu tab 'OPR' ({len(raw_values)} dòng)")
    except Exception as e:
        print(f"❌ Lỗi mở hoặc đọc tab 'OPR': {e}")
        sys.exit(1)

    if len(raw_values) < 2:
        print("❌ Tab data 'OPR' không có data hoặc chỉ có header!")
        sys.exit(0)

    headers = raw_values[0]
    
    # 1.1 Determine Target Date (N) and Previous Date (N-1)
    tz = pytz.timezone("Asia/Ho_Chi_Minh")
    now = datetime.now(tz)
    target_date = None
    if len(sys.argv) > 1:
        try:
            target_date = datetime.strptime(sys.argv[1], "%Y-%m-%d")
            print(f"👉 Chạy báo cáo cho ngày được chỉ định: {sys.argv[1]}")
        except ValueError:
            print(f"⚠️ Định dạng ngày chỉ định không đúng (%Y-%m-%d).")
            
    if not target_date:
        target_date = now - timedelta(days=1)
        print(f"👉 Chạy báo cáo cho ngày N: {target_date.strftime('%Y-%m-%d')}")
        
    prev_date = target_date - timedelta(days=1)
    
    target_date_str = target_date.strftime("%Y-%m-%d")
    prev_date_str = prev_date.strftime("%Y-%m-%d")
    
    date_n_label = target_date.strftime('%d/%m')
    date_n1_label = prev_date.strftime('%d/%m')

    print(f"📅 Ngày N: {target_date_str} ({date_n_label}) | Ngày N-1: {prev_date_str} ({date_n1_label})")

    col_date = find_col(headers, ["ngayltc", "ngay_ltc", "ngày ltc"], 0)
    col_am = find_col(headers, ["am"], 10)
    col_create_time = find_col(headers, ["khung_gio_tao", "khung giờ tạo", "khung gio tao"], 1)
    col_vol = find_col(headers, ["don_ltc", "vol_ltc", "don ltc"], 8)
    col_ot = find_col(headers, ["don_ontime", "ot", "don ontime"], 9)
    col_late = find_col(headers, ["đơn trễ", "don tre", "don_tre"], -1)

    print(f"-> Mapped columns: Date={col_date}, AM={col_am}, KhungGiờTạo={col_create_time}, Vol={col_vol}, OT={col_ot}, Late={col_late}")

    # 2. Process Data for Date N & Date N-1
    am_map_all = {}
    am_map_9to19 = {}
    am_map_19to9 = {}

    prev_am_map_all = {}
    prev_am_map_9to19 = {}
    prev_am_map_19to9 = {}

    for row_idx, row in enumerate(raw_values[1:], start=2):
        valid_indices = [col_date, col_am, col_create_time, col_vol, col_ot]
        if col_late != -1:
            valid_indices.append(col_late)
        if len(row) <= max(valid_indices):
            continue
            
        row_date_str = str(row[col_date]).strip()
        is_target = (row_date_str == target_date_str)
        is_prev = (row_date_str == prev_date_str)

        if not (is_target or is_prev):
            continue
        
        am_raw = row[col_am]
        if not am_raw:
            continue
        am_name = normalize_am_name(am_raw)
        if am_name == "Không xác định":
            continue
            
        try:
            vol = float(row[col_vol])
        except (ValueError, IndexError):
            vol = 0.0
        try:
            ot = float(row[col_ot])
        except (ValueError, IndexError):
            ot = 0.0
            
        if col_late != -1 and col_late < len(row):
            try:
                late = float(row[col_late])
            except ValueError:
                late = max(0.0, vol - ot)
        else:
            late = max(0.0, vol - ot)
            
        time_raw = str(row[col_create_time]).strip()
        win = ""
        if "9h-19h" in time_raw:
            win = "1.Tạo từ 9h-19h"
        elif "19h-9h" in time_raw or "trước 9h" in time_raw.lower() or "sau 19h" in time_raw.lower() or "truoc 9h" in time_raw.lower() or "sau 19h" in time_raw.lower():
            win = "2.Tạo từ 19h-9h"

        if is_target:
            # Data N
            if am_name not in am_map_all:
                am_map_all[am_name] = {"vol": 0, "ot": 0, "late": 0}
            am_map_all[am_name]["vol"] += vol
            am_map_all[am_name]["ot"] += ot
            am_map_all[am_name]["late"] += late
            
            if win == "1.Tạo từ 9h-19h":
                if am_name not in am_map_9to19:
                    am_map_9to19[am_name] = {"vol": 0, "ot": 0, "late": 0}
                am_map_9to19[am_name]["vol"] += vol
                am_map_9to19[am_name]["ot"] += ot
                am_map_9to19[am_name]["late"] += late
                
            if win == "2.Tạo từ 19h-9h":
                if am_name not in am_map_19to9:
                    am_map_19to9[am_name] = {"vol": 0, "ot": 0, "late": 0}
                am_map_19to9[am_name]["vol"] += vol
                am_map_19to9[am_name]["ot"] += ot
                am_map_19to9[am_name]["late"] += late

        elif is_prev:
            # Data N-1
            if am_name not in prev_am_map_all:
                prev_am_map_all[am_name] = {"vol": 0, "ot": 0, "late": 0}
            prev_am_map_all[am_name]["vol"] += vol
            prev_am_map_all[am_name]["ot"] += ot
            prev_am_map_all[am_name]["late"] += late
            
            if win == "1.Tạo từ 9h-19h":
                if am_name not in prev_am_map_9to19:
                    prev_am_map_9to19[am_name] = {"vol": 0, "ot": 0, "late": 0}
                prev_am_map_9to19[am_name]["vol"] += vol
                prev_am_map_9to19[am_name]["ot"] += ot
                prev_am_map_9to19[am_name]["late"] += late
                
            if win == "2.Tạo từ 19h-9h":
                if am_name not in prev_am_map_19to9:
                    prev_am_map_19to9[am_name] = {"vol": 0, "ot": 0, "late": 0}
                prev_am_map_19to9[am_name]["vol"] += vol
                prev_am_map_19to9[am_name]["ot"] += ot
                prev_am_map_19to9[am_name]["late"] += late

    # Get/Create Worksheet REPORT_OPR
    try:
        report_sheet = sh.worksheet("REPORT_OPR")
        report_sheet.clear()
        print("✔️ Đã clear sheet cũ 'REPORT_OPR'")
    except gspread.exceptions.WorksheetNotFound:
        report_sheet = sh.add_worksheet(title="REPORT_OPR", rows="1000", cols="20")
        print("✔️ Đã tạo sheet mới 'REPORT_OPR'")

    # Set hidden gridlines
    try:
        sh.batch_update({
            "requests": [
                {
                    "updateSheetProperties": {
                        "properties": {
                            "sheetId": report_sheet.id,
                            "gridProperties": {
                                "hideGridlines": True
                            }
                        },
                        "fields": "gridProperties.hideGridlines"
                    }
                }
            ]
        })
    except Exception as e:
        print(f"⚠️ Không thể ẩn gridlines: {e}")

    # Preparing grid variables
    grid_rows = []
    requests = []
    
    # Clear all formatting & unmerge all ranges first
    requests.append({
        "repeatCell": {
            "range": {
                "sheetId": report_sheet.id,
                "startRowIndex": 0,
                "endRowIndex": 1000,
                "startColumnIndex": 0,
                "endColumnIndex": 20
            },
            "cell": {
                "userEnteredFormat": {}
            },
            "fields": "userEnteredFormat"
        }
    })
    requests.append({
        "unmergeCells": {
            "range": {
                "sheetId": report_sheet.id,
                "startRowIndex": 0,
                "endRowIndex": 1000,
                "startColumnIndex": 0,
                "endColumnIndex": 20
            }
        }
    })
    
    time_label = now.strftime('%H:%M')
    date_label = target_date.strftime('%d/%m/%Y')

    # Main Titles
    grid_rows.append(["BÁO CÁO HIỆU SUẤT VẬN HÀNH OPR", "", "", "", "", "", "", ""])
    requests.append(merge_request(report_sheet.id, 0, 1, 0, 8))
    requests.append(cell_format_request(report_sheet.id, 0, 1, 0, 8, {
        "textFormat": {"fontSize": 16, "bold": True, "foregroundColor": make_color("#1E293B")},
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "MIDDLE"
    }))
    requests.append(row_height_request(report_sheet.id, 0, 1, 35))

    grid_rows.append([f"Mốc cập nhật: {time_label} ngày {date_label} | So sánh với ngày {date_n1_label}", "", "", "", "", "", "", ""])
    requests.append(merge_request(report_sheet.id, 1, 2, 0, 8))
    requests.append(cell_format_request(report_sheet.id, 1, 2, 0, 8, {
        "textFormat": {"fontSize": 10, "italic": True, "foregroundColor": make_color("#475569")},
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "MIDDLE"
    }))
    requests.append(row_height_request(report_sheet.id, 1, 2, 22))

    grid_rows.append(["", "", "", "", "", "", "", ""])
    grid_rows.append(["", "", "", "", "", "", "", ""])

    cur_row = 4 # 0-indexed index for row 5

    def write_opr_table_py(title, data_map, prev_data_map, title_bg, header_bg, alt_row_color):
        nonlocal cur_row
        start_row = cur_row
        
        # 1. Title Row
        grid_rows.append([title, "", "", "", "", "", "", ""])
        requests.append(merge_request(report_sheet.id, cur_row, cur_row + 1, 0, 8))
        requests.append(cell_format_request(report_sheet.id, cur_row, cur_row + 1, 0, 8, {
            "backgroundColor": make_color(title_bg),
            "textFormat": {"fontSize": 12, "bold": True, "foregroundColor": make_color("#FFFFFF")},
            "horizontalAlignment": "CENTER",
            "verticalAlignment": "MIDDLE"
        }))
        requests.append(row_height_request(report_sheet.id, cur_row, cur_row + 1, 25))
        cur_row += 1

        # 2. Header Row (8 columns)
        col_headers = ["AM", "Tổng đơn", "Đơn đúng hạn", "Số đơn trễ", "Tỷ trọng lỗi trễ", f"%OPR ({date_n_label})", f"%OPR ({date_n1_label})", "vs N-1"]
        grid_rows.append(col_headers)
        requests.append(cell_format_request(report_sheet.id, cur_row, cur_row + 1, 0, 8, {
            "backgroundColor": make_color(header_bg),
            "textFormat": {"fontSize": 10, "bold": True, "foregroundColor": make_color("#FFFFFF")},
            "horizontalAlignment": "CENTER",
            "verticalAlignment": "MIDDLE"
        }))
        requests.append(row_height_request(report_sheet.id, cur_row, cur_row + 1, 25))
        cur_row += 1

        # Calculate totals Date N
        am_list = []
        total_late = 0
        total_vol = 0
        total_ot = 0
        for am, d in data_map.items():
            am_list.append(am)
            total_late += d["late"]
            total_vol += d["vol"]
            total_ot += d["ot"]

        # Calculate totals Date N-1
        prev_total_vol = sum(d["vol"] for d in prev_data_map.values())
        prev_total_ot = sum(d["ot"] for d in prev_data_map.values())

        # Sort: vol desc, late desc
        am_list.sort(key=lambda item: (-data_map[item]["vol"], -data_map[item]["late"]))

        # 3. Body Rows
        for idx, am in enumerate(am_list):
            d = data_map[am]
            prev_d = prev_data_map.get(am, {"vol": 0, "ot": 0, "late": 0})

            late_share = d["late"] / total_late if total_late > 0 else 0.0
            opr_val = d["ot"] / d["vol"] if d["vol"] > 0 else 0.0
            prev_opr_val = prev_d["ot"] / prev_d["vol"] if prev_d["vol"] > 0 else 0.0
            
            diff_opr = opr_val - prev_opr_val

            grid_rows.append([am, int(d["vol"]), int(d["ot"]), int(d["late"]), late_share, opr_val, prev_opr_val, diff_opr])
            
            # Formats for body columns A-E
            color_row = "#FFFFFF" if idx % 2 == 0 else alt_row_color
            requests.append(cell_format_request(report_sheet.id, cur_row, cur_row + 1, 0, 5, {
                "backgroundColor": make_color(color_row)
            }))
            
            # Format %OPR Date N cell in col F (index 5)
            if opr_val >= 0.80:
                requests.append(cell_format_request(report_sheet.id, cur_row, cur_row + 1, 5, 6, {
                    "backgroundColor": make_color("#E2EFDA"),
                    "textFormat": {"foregroundColor": make_color("#375623"), "bold": True}
                }))
            else:
                requests.append(cell_format_request(report_sheet.id, cur_row, cur_row + 1, 5, 6, {
                    "backgroundColor": make_color("#FCE4D6"),
                    "textFormat": {"foregroundColor": make_color("#C65911"), "bold": True}
                }))

            # Format %OPR Date N-1 cell in col G (index 6)
            if prev_opr_val >= 0.80:
                requests.append(cell_format_request(report_sheet.id, cur_row, cur_row + 1, 6, 7, {
                    "backgroundColor": make_color("#E2EFDA"),
                    "textFormat": {"foregroundColor": make_color("#375623")}
                }))
            else:
                requests.append(cell_format_request(report_sheet.id, cur_row, cur_row + 1, 6, 7, {
                    "backgroundColor": make_color("#FCE4D6"),
                    "textFormat": {"foregroundColor": make_color("#C65911")}
                }))

            # Format vs N-1 cell in col H (index 7)
            if diff_opr > 0:
                requests.append(cell_format_request(report_sheet.id, cur_row, cur_row + 1, 7, 8, {
                    "backgroundColor": make_color(color_row),
                    "textFormat": {"foregroundColor": make_color("#16A34A"), "bold": True}
                }))
            elif diff_opr < 0:
                requests.append(cell_format_request(report_sheet.id, cur_row, cur_row + 1, 7, 8, {
                    "backgroundColor": make_color(color_row),
                    "textFormat": {"foregroundColor": make_color("#DC2626"), "bold": True}
                }))
            else:
                requests.append(cell_format_request(report_sheet.id, cur_row, cur_row + 1, 7, 8, {
                    "backgroundColor": make_color(color_row),
                    "textFormat": {"foregroundColor": make_color("#475569")}
                }))

            # Base format for row
            requests.append(cell_format_request(report_sheet.id, cur_row, cur_row + 1, 0, 8, {
                "textFormat": {"fontSize": 9},
                "verticalAlignment": "MIDDLE",
                "horizontalAlignment": "CENTER"
            }))
            
            # Left align AM name
            requests.append(cell_format_request(report_sheet.id, cur_row, cur_row + 1, 0, 1, {
                "horizontalAlignment": "LEFT"
            }))
            
            # Number formats
            requests.append(cell_format_request(report_sheet.id, cur_row, cur_row + 1, 1, 4, {
                "numberFormat": {"type": "NUMBER", "pattern": "#,##0"}
            }))
            requests.append(cell_format_request(report_sheet.id, cur_row, cur_row + 1, 4, 7, {
                "numberFormat": {"type": "PERCENT", "pattern": "0.00%"}
            }))
            requests.append(cell_format_request(report_sheet.id, cur_row, cur_row + 1, 7, 8, {
                "numberFormat": {"type": "PERCENT", "pattern": "+0.00%;-0.00%;0.00%"}
            }))

            requests.append(row_height_request(report_sheet.id, cur_row, cur_row + 1, 20))
            cur_row += 1

        # 4. Grand Total Row
        total_opr = total_ot / total_vol if total_vol > 0 else 0.0
        prev_total_opr = prev_total_ot / prev_total_vol if prev_total_vol > 0 else 0.0
        diff_total_opr = total_opr - prev_total_opr

        grid_rows.append(["Grand Total", int(total_vol), int(total_ot), int(total_late), 1.0, total_opr, prev_total_opr, diff_total_opr])
        
        requests.append(cell_format_request(report_sheet.id, cur_row, cur_row + 1, 0, 5, {
            "backgroundColor": make_color("#D9D9D9")
        }))
        
        # Grand Total %OPR N
        if total_opr >= 0.80:
            requests.append(cell_format_request(report_sheet.id, cur_row, cur_row + 1, 5, 6, {
                "backgroundColor": make_color("#E2EFDA"),
                "textFormat": {"foregroundColor": make_color("#375623"), "bold": True}
            }))
        else:
            requests.append(cell_format_request(report_sheet.id, cur_row, cur_row + 1, 5, 6, {
                "backgroundColor": make_color("#FCE4D6"),
                "textFormat": {"foregroundColor": make_color("#C65911"), "bold": True}
            }))

        # Grand Total %OPR N-1
        if prev_total_opr >= 0.80:
            requests.append(cell_format_request(report_sheet.id, cur_row, cur_row + 1, 6, 7, {
                "backgroundColor": make_color("#E2EFDA"),
                "textFormat": {"foregroundColor": make_color("#375623"), "bold": True}
            }))
        else:
            requests.append(cell_format_request(report_sheet.id, cur_row, cur_row + 1, 6, 7, {
                "backgroundColor": make_color("#FCE4D6"),
                "textFormat": {"foregroundColor": make_color("#C65911"), "bold": True}
            }))

        # Grand Total vs N-1
        diff_color = "#16A34A" if diff_total_opr > 0 else ("#DC2626" if diff_total_opr < 0 else "#475569")
        requests.append(cell_format_request(report_sheet.id, cur_row, cur_row + 1, 7, 8, {
            "backgroundColor": make_color("#D9D9D9"),
            "textFormat": {"foregroundColor": make_color(diff_color), "bold": True}
        }))

        requests.append(cell_format_request(report_sheet.id, cur_row, cur_row + 1, 0, 8, {
            "textFormat": {"fontSize": 10, "bold": True},
            "verticalAlignment": "MIDDLE",
            "horizontalAlignment": "CENTER"
        }))
        
        requests.append(cell_format_request(report_sheet.id, cur_row, cur_row + 1, 0, 1, {
            "horizontalAlignment": "LEFT"
        }))
        
        # Number formats for Grand Total
        requests.append(cell_format_request(report_sheet.id, cur_row, cur_row + 1, 1, 4, {
            "numberFormat": {"type": "NUMBER", "pattern": "#,##0"}
        }))
        requests.append(cell_format_request(report_sheet.id, cur_row, cur_row + 1, 4, 7, {
            "numberFormat": {"type": "PERCENT", "pattern": "0.00%"}
        }))
        requests.append(cell_format_request(report_sheet.id, cur_row, cur_row + 1, 7, 8, {
            "numberFormat": {"type": "PERCENT", "pattern": "+0.00%;-0.00%;0.00%"}
        }))

        requests.append(row_height_request(report_sheet.id, cur_row, cur_row + 1, 22))
        
        # Border
        requests.append(border_request(report_sheet.id, start_row + 1, cur_row + 1, 0, 8))
        
        cur_row += 1
        
        # Blank rows spacing (3 rows)
        grid_rows.append(["", "", "", "", "", "", "", ""])
        grid_rows.append(["", "", "", "", "", "", "", ""])
        grid_rows.append(["", "", "", "", "", "", "", ""])
        cur_row += 3

        return {
            "start_row": start_row,
            "end_row": cur_row - 3,
            "totalVol": total_vol,
            "totalOt": total_ot,
            "totalLate": total_late,
            "totalOpr": total_opr,
            "prevTotalOpr": prev_total_opr,
            "sortedAmList": am_list
        }

    # Bảng 1, 2, 3
    r1 = write_opr_table_py("%OPR TTS THEO AM - KPI 80%",    am_map_all,   prev_am_map_all,   "#1E293B", "#475569", "#F8FAFC")
    r2 = write_opr_table_py("%OPR TTS tạo từ 9h-19h",        am_map_9to19, prev_am_map_9to19, "#047857", "#10B981", "#F0FDF4")
    r3 = write_opr_table_py("%OPR TTS tạo từ 19h-9h",        am_map_19to9, prev_am_map_19to9, "#4338CA", "#6366F1", "#EEF2F6")

    # === BẢNG 4: TỶ TRỌNG LỖI OPR TTS ===
    t4StartRow = cur_row
    
    # Title row
    grid_rows.append(["Tỷ trọng lỗi OPR TTS theo khung giờ tạo", "", "", "", "", "", "", ""])
    requests.append(merge_request(report_sheet.id, cur_row, cur_row + 1, 0, 8))
    requests.append(cell_format_request(report_sheet.id, cur_row, cur_row + 1, 0, 8, {
        "backgroundColor": make_color("#C2410C"),
        "textFormat": {"fontSize": 12, "bold": True, "foregroundColor": make_color("#FFFFFF")},
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "MIDDLE"
    }))
    requests.append(row_height_request(report_sheet.id, cur_row, cur_row + 1, 25))
    cur_row += 1

    # Header row 1
    grid_rows.append(["Số đơn trễ", "Khung giờ tạo", "", "", "", "", "", ""])
    requests.append(merge_request(report_sheet.id, cur_row, cur_row + 1, 0, 1))
    requests.append(cell_format_request(report_sheet.id, cur_row, cur_row + 1, 0, 1, {
        "backgroundColor": make_color("#64748B"),
        "textFormat": {"fontSize": 10, "bold": True, "foregroundColor": make_color("#FFFFFF")},
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "MIDDLE"
    }))
    requests.append(merge_request(report_sheet.id, cur_row, cur_row + 1, 1, 8))
    requests.append(cell_format_request(report_sheet.id, cur_row, cur_row + 1, 1, 8, {
        "backgroundColor": make_color("#EA580C"),
        "textFormat": {"fontSize": 10, "bold": True, "foregroundColor": make_color("#FFFFFF")},
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "MIDDLE"
    }))
    requests.append(row_height_request(report_sheet.id, cur_row, cur_row + 1, 22))
    cur_row += 1

    # Header row 2
    grid_rows.append(["AM", "1.Tạo từ 9h-19h", "", "", "", "2.Tạo từ 19h-9h", "", ""])
    requests.append(cell_format_request(report_sheet.id, cur_row, cur_row + 1, 0, 1, {
        "backgroundColor": make_color("#64748B"),
        "textFormat": {"fontSize": 10, "bold": True, "foregroundColor": make_color("#FFFFFF")},
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "MIDDLE"
    }))
    requests.append(merge_request(report_sheet.id, cur_row, cur_row + 1, 1, 5))
    requests.append(cell_format_request(report_sheet.id, cur_row, cur_row + 1, 1, 5, {
        "backgroundColor": make_color("#2563EB"),
        "textFormat": {"fontSize": 10, "bold": True, "foregroundColor": make_color("#FFFFFF")},
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "MIDDLE"
    }))
    requests.append(merge_request(report_sheet.id, cur_row, cur_row + 1, 5, 8))
    requests.append(cell_format_request(report_sheet.id, cur_row, cur_row + 1, 5, 8, {
        "backgroundColor": make_color("#D97706"),
        "textFormat": {"fontSize": 10, "bold": True, "foregroundColor": make_color("#FFFFFF")},
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "MIDDLE"
    }))
    requests.append(row_height_request(report_sheet.id, cur_row, cur_row + 1, 22))
    cur_row += 1

    # Calculate filtered data
    filtered_AMs = []
    total_late_9to19 = 0
    total_late_19to9 = 0
    total_late_all = 0

    for am in r1["sortedAmList"]:
        d_all = am_map_all.get(am, {"late": 0.0})
        if d_all["late"] > 0:
            filtered_AMs.append(am)
            d9 = am_map_9to19.get(am, {"late": 0.0})
            d19 = am_map_19to9.get(am, {"late": 0.0})
            total_late_9to19 += d9["late"]
            total_late_19to9 += d19["late"]
            total_late_all += d_all["late"]

    # Body Table 4
    for idx, am in enumerate(filtered_AMs):
        d_all = am_map_all[am]
        d9 = am_map_9to19.get(am, {"late": 0.0})
        d19 = am_map_19to9.get(am, {"late": 0.0})
        
        val9 = d9["late"] / d_all["late"] if d9["late"] > 0 else ""
        val19 = d19["late"] / d_all["late"] if d19["late"] > 0 else ""

        grid_rows.append([am, val9, "", "", "", val19, "", ""])
        
        color_row = "#FFFFFF" if idx % 2 == 0 else "#FFFDF5"
        requests.append(cell_format_request(report_sheet.id, cur_row, cur_row + 1, 0, 8, {
            "backgroundColor": make_color(color_row),
            "textFormat": {"fontSize": 9},
            "verticalAlignment": "MIDDLE",
            "horizontalAlignment": "CENTER"
        }))
        
        # Merges for body cols
        requests.append(merge_request(report_sheet.id, cur_row, cur_row + 1, 1, 5))
        requests.append(merge_request(report_sheet.id, cur_row, cur_row + 1, 5, 8))

        requests.append(cell_format_request(report_sheet.id, cur_row, cur_row + 1, 0, 1, {
            "horizontalAlignment": "LEFT"
        }))
        
        # Number format %
        requests.append(cell_format_request(report_sheet.id, cur_row, cur_row + 1, 1, 8, {
            "numberFormat": {"type": "PERCENT", "pattern": "0.00%"}
        }))

        requests.append(row_height_request(report_sheet.id, cur_row, cur_row + 1, 20))
        cur_row += 1

    # Grand Total Table 4
    gt_val9 = total_late_9to19 / total_late_all if total_late_all > 0 else 0.0
    gt_val19 = total_late_19to9 / total_late_all if total_late_all > 0 else 0.0
    grid_rows.append(["Grand Total", gt_val9, "", "", "", gt_val19, "", ""])
    
    requests.append(cell_format_request(report_sheet.id, cur_row, cur_row + 1, 0, 8, {
        "backgroundColor": make_color("#D9D9D9"),
        "textFormat": {"fontSize": 10, "bold": True},
        "verticalAlignment": "MIDDLE",
        "horizontalAlignment": "CENTER"
    }))
    
    requests.append(merge_request(report_sheet.id, cur_row, cur_row + 1, 1, 5))
    requests.append(merge_request(report_sheet.id, cur_row, cur_row + 1, 5, 8))
    
    requests.append(cell_format_request(report_sheet.id, cur_row, cur_row + 1, 0, 1, {
        "horizontalAlignment": "LEFT"
    }))
    requests.append(cell_format_request(report_sheet.id, cur_row, cur_row + 1, 1, 8, {
        "numberFormat": {"type": "PERCENT", "pattern": "0.00%"}
    }))
    requests.append(row_height_request(report_sheet.id, cur_row, cur_row + 1, 22))

    # Border Table 4
    requests.append(border_request(report_sheet.id, t4StartRow + 1, cur_row + 1, 0, 8))
    
    r4 = {
        "start_row": t4StartRow,
        "end_row": cur_row,
        "totalLateAll": total_late_all,
        "totalLate9to19": total_late_9to19,
        "totalLate19to9": total_late_19to9
    }
    
    cur_row += 1

    # Apply font Arial globally to the written range in Google Sheets
    requests.append({
        "repeatCell": {
            "range": {
                "sheetId": report_sheet.id,
                "startRowIndex": 0,
                "endRowIndex": cur_row,
                "startColumnIndex": 0,
                "endColumnIndex": 8
            },
            "cell": {
                "userEnteredFormat": {
                    "textFormat": {"fontFamily": "Arial"}
                }
            },
            "fields": "userEnteredFormat.textFormat.fontFamily"
        }
    })

    # Set Column widths
    col_widths = [160, 85, 95, 85, 95, 85, 85, 80]
    for c_idx, width in enumerate(col_widths):
        requests.append(col_width_request(report_sheet.id, c_idx, c_idx + 1, width))

    # Writing values in a single call
    print("-> Đang cập nhật dữ liệu hàng loạt lên Google Sheets...")
    filled_grid_rows = []
    for r in grid_rows:
        if len(r) < 8:
            r = r + [""] * (8 - len(r))
        filled_grid_rows.append(r)

    report_sheet.update(range_name=f"A1:H{len(filled_grid_rows)}", values=filled_grid_rows, value_input_option="USER_ENTERED")
    print("✔️ Đã cập nhật xong dữ liệu.")

    # Apply formatting requests
    print("-> Đang áp dụng định dạng (màu sắc, border, merge)...")
    try:
        sh.batch_update({"requests": requests})
        print("✔️ Đã định dạng bảng thành công trên Google Sheets.")
    except Exception as e:
        print(f"❌ Lỗi định dạng Google Sheets: {e}")
        sys.exit(1)

    # 3. HTML Rendering using Local HTML & Playwright Element Screenshot
    print("-> Đang build HTML với phong cách thiết kế hiện đại (Modern Premium Visuals)...")
    
    def build_table_html(title, am_list, data_map, prev_data_map, header_gradient, is_table_4=False):
        if not is_table_4:
            total_vol = sum(data_map[am]["vol"] for am in am_list)
            total_ot = sum(data_map[am]["ot"] for am in am_list)
            total_late = sum(data_map[am]["late"] for am in am_list)
            total_opr = total_ot / total_vol if total_vol > 0 else 0.0

            prev_total_vol = sum(prev_data_map.get(am, {"vol": 0})["vol"] for am in am_list)
            prev_total_ot = sum(prev_data_map.get(am, {"ot": 0})["ot"] for am in am_list)
            prev_total_opr = prev_total_ot / prev_total_vol if prev_total_vol > 0 else 0.0

            diff_total_opr = total_opr - prev_total_opr
            
            body_rows = ""
            for idx, am in enumerate(am_list):
                d = data_map[am]
                prev_d = prev_data_map.get(am, {"vol": 0, "ot": 0, "late": 0})

                late_share = d["late"] / total_late if total_late > 0 else 0.0
                opr_val = d["ot"] / d["vol"] if d["vol"] > 0 else 0.0
                prev_opr_val = prev_d["ot"] / prev_d["vol"] if prev_d["vol"] > 0 else 0.0
                diff_opr = opr_val - prev_opr_val
                
                color_row = "#FFFFFF" if idx % 2 == 0 else "#F8FAFC"
                opr_badge = f'<span class="badge badge-success">{opr_val * 100:.2f}%</span>' if opr_val >= 0.80 else f'<span class="badge badge-danger">{opr_val * 100:.2f}%</span>'
                prev_opr_badge = f'<span class="badge badge-success">{prev_opr_val * 100:.2f}%</span>' if prev_opr_val >= 0.80 else f'<span class="badge badge-danger">{prev_opr_val * 100:.2f}%</span>'
                
                if diff_opr > 0:
                    diff_badge = f'<span class="badge badge-up">▲ +{diff_opr * 100:.2f}%</span>'
                elif diff_opr < 0:
                    diff_badge = f'<span class="badge badge-down">▼ {diff_opr * 100:.2f}%</span>'
                else:
                    diff_badge = f'<span class="badge badge-neutral">● 0.00%</span>'

                body_rows += f"""
                <tr style="background-color: {color_row};">
                    <td style="text-align: left; font-weight: 600; color: #1e293b;">{am}</td>
                    <td class="number-cell">{int(d['vol']):,}</td>
                    <td class="number-cell" style="color: #166534; font-weight: 600;">{int(d['ot']):,}</td>
                    <td class="number-cell" style="color: #991b1b; font-weight: 600;">{int(d['late']):,}</td>
                    <td class="number-cell">{late_share * 100:.2f}%</td>
                    <td>{opr_badge}</td>
                    <td>{prev_opr_badge}</td>
                    <td>{diff_badge}</td>
                </tr>
                """
            
            opr_badge_gt = f'<span class="badge badge-success">{total_opr * 100:.2f}%</span>' if total_opr >= 0.80 else f'<span class="badge badge-danger">{total_opr * 100:.2f}%</span>'
            prev_opr_badge_gt = f'<span class="badge badge-success">{prev_total_opr * 100:.2f}%</span>' if prev_total_opr >= 0.80 else f'<span class="badge badge-danger">{prev_total_opr * 100:.2f}%</span>'
            
            if diff_total_opr > 0:
                diff_gt_badge = f'<span class="badge badge-up">▲ +{diff_total_opr * 100:.2f}%</span>'
            elif diff_total_opr < 0:
                diff_gt_badge = f'<span class="badge badge-down">▼ {diff_total_opr * 100:.2f}%</span>'
            else:
                diff_gt_badge = f'<span class="badge badge-neutral">● 0.00%</span>'

            grand_total_row = f"""
            <tr class="grand-total-row">
                <td style="text-align: left;">Grand Total</td>
                <td class="number-cell">{int(total_vol):,}</td>
                <td class="number-cell" style="color: #166534;">{int(total_ot):,}</td>
                <td class="number-cell" style="color: #991b1b;">{int(total_late):,}</td>
                <td class="number-cell">100.00%</td>
                <td>{opr_badge_gt}</td>
                <td>{prev_opr_badge_gt}</td>
                <td>{diff_gt_badge}</td>
            </tr>
            """
            
            html = f"""
            <div class="table-card">
              <table>
                  <thead>
                    <tr class="title-row" style="background: {header_gradient};">
                        <th colspan="8">{title}</th>
                    </tr>
                    <tr class="header-row" style="background: #334155;">
                        <th>AM</th>
                        <th>Tổng đơn</th>
                        <th>Đơn đúng hạn</th>
                        <th>Số đơn trễ</th>
                        <th>Tỷ trọng lỗi</th>
                        <th>%OPR ({date_n_label})</th>
                        <th>%OPR ({date_n1_label})</th>
                        <th>vs N-1</th>
                    </tr>
                  </thead>
                  <tbody>
                    {body_rows}
                  </tbody>
                  <tfoot>
                    {grand_total_row}
                  </tfoot>
              </table>
            </div>
            """
            return html
        else:
            total_late_all = sum(am_map_all[am]["late"] for am in am_list)
            total_late_9to19 = sum(am_map_9to19.get(am, {"late": 0.0})["late"] for am in am_list)
            total_late_19to9 = sum(am_map_19to9.get(am, {"late": 0.0})["late"] for am in am_list)
            
            body_rows = ""
            for idx, am in enumerate(am_list):
                d_all = am_map_all[am]
                d9 = am_map_9to19.get(am, {"late": 0.0})
                d19 = am_map_19to9.get(am, {"late": 0.0})
                
                val9_pct = f"{d9['late'] / d_all['late'] * 100:.2f}%" if d9['late'] > 0 else "-"
                val19_pct = f"{d19['late'] / d_all['late'] * 100:.2f}%" if d19['late'] > 0 else "-"
                
                color_row = "#FFFFFF" if idx % 2 == 0 else "#FFFDF5"
                
                body_rows += f"""
                <tr style="background-color: {color_row};">
                    <td style="text-align: left; font-weight: 600; color: #1e293b;">{am}</td>
                    <td colspan="4" class="number-cell" style="color: #2563eb; font-weight: 600;">{val9_pct}</td>
                    <td colspan="3" class="number-cell" style="color: #d97706; font-weight: 600;">{val19_pct}</td>
                </tr>
                """
                
            gt_val9 = f"{total_late_9to19 / total_late_all * 100:.2f}%" if total_late_all > 0 else "0.00%"
            gt_val19 = f"{total_late_19to9 / total_late_all * 100:.2f}%" if total_late_all > 0 else "0.00%"
            
            grand_total_row = f"""
            <tr class="grand-total-row">
                <td style="text-align: left;">Grand Total</td>
                <td colspan="4" class="number-cell" style="color: #2563eb;">{gt_val9}</td>
                <td colspan="3" class="number-cell" style="color: #d97706;">{gt_val19}</td>
            </tr>
            """
            
            html = f"""
            <div class="table-card">
              <table class="table4">
                  <thead>
                    <tr class="title-row" style="background: linear-gradient(135deg, #c2410c, #ea580c);">
                        <th colspan="8">{title}</th>
                    </tr>
                    <tr style="background-color: #475569; color: #FFFFFF; font-weight: 600; text-align: center; height: 26px;">
                        <td style="border: 1px solid #cbd5e1; width: 160px; font-weight: 700;">Số đơn trễ</td>
                        <td colspan="7" style="border: 1px solid #cbd5e1; background-color: #ea580C; font-weight: 700; width: 660px;">Khung giờ tạo</td>
                    </tr>
                    <tr style="background-color: #475569; color: #FFFFFF; font-weight: 600; text-align: center; height: 26px;">
                        <td style="border: 1px solid #cbd5e1; font-weight: 700;">AM</td>
                        <td colspan="4" style="border: 1px solid #cbd5e1; background-color: #2563eb; font-weight: 700; width: 380px;">1. Tạo từ 9h-19h</td>
                        <td colspan="3" style="border: 1px solid #cbd5e1; background-color: #d97706; font-weight: 700; width: 280px;">2. Tạo từ 19h-9h</td>
                    </tr>
                  </thead>
                  <tbody>
                    {body_rows}
                  </tbody>
                  <tfoot>
                    {grand_total_row}
                  </tfoot>
              </table>
            </div>
            """
            return html

    t1_html = build_table_html("%OPR TTS THEO AM - KPI 80%", r1["sortedAmList"], am_map_all, prev_am_map_all, "linear-gradient(135deg, #1e293b, #334155)")
    t2_html = build_table_html("%OPR TTS tạo từ 9h-19h", r2["sortedAmList"], am_map_9to19, prev_am_map_9to19, "linear-gradient(135deg, #065f46, #047857)")
    t3_html = build_table_html("%OPR TTS tạo từ 19h-9h", r3["sortedAmList"], am_map_19to9, prev_am_map_19to9, "linear-gradient(135deg, #3730a3, #4338ca)")
    t4_html = build_table_html("Tỷ trọng lỗi OPR TTS theo khung giờ tạo", filtered_AMs, None, None, "linear-gradient(135deg, #c2410c, #ea580c)", is_table_4=True)

    full_html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
  
  body {{
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    margin: 0;
    padding: 24px;
    background-color: #f8fafc;
    color: #0f172a;
    -webkit-font-smoothing: antialiased;
  }}
  .table-card {{
    background: #ffffff;
    border-radius: 12px;
    box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.08), 0 4px 6px -2px rgba(15, 23, 42, 0.03);
    border: 1px solid #e2e8f0;
    overflow: hidden;
    display: inline-block;
    margin-bottom: 24px;
  }}
  table {{
    border-collapse: collapse;
    width: 820px;
    font-size: 13px;
  }}
  .title-row th {{
    font-size: 15px;
    font-weight: 700;
    color: #ffffff;
    text-align: center;
    padding: 12px 16px;
    letter-spacing: 0.5px;
    text-transform: uppercase;
  }}
  .header-row th {{
    font-size: 12px;
    font-weight: 600;
    color: #ffffff;
    text-align: center;
    padding: 10px 8px;
    letter-spacing: 0.3px;
    border-bottom: 2px solid rgba(255,255,255,0.1);
  }}
  th, td {{
    padding: 8px 10px;
    vertical-align: middle;
    border-bottom: 1px solid #f1f5f9;
    text-align: center;
  }}
  tr:last-child td {{
    border-bottom: none;
  }}
  
  /* Pillar Badges */
  .badge {{
    display: inline-block;
    padding: 3px 9px;
    border-radius: 20px;
    font-size: 11.5px;
    font-weight: 700;
    text-align: center;
    min-width: 58px;
  }}
  .badge-success {{
    background-color: #dcfce7;
    color: #15803d;
    border: 1px solid #bbf7d0;
  }}
  .badge-danger {{
    background-color: #fee2e2;
    color: #b91c1c;
    border: 1px solid #fecaca;
  }}
  .badge-up {{
    background-color: #f0fdf4;
    color: #166534;
    border: 1px solid #bbf7d0;
  }}
  .badge-down {{
    background-color: #fef2f2;
    color: #991b1b;
    border: 1px solid #fecaca;
  }}
  .badge-neutral {{
    background-color: #f1f5f9;
    color: #475569;
    border: 1px solid #e2e8f0;
  }}

  .number-cell {{
    font-family: 'Consolas', 'Menlo', monospace;
    font-size: 12.5px;
    color: #334155;
  }}
  
  .grand-total-row {{
    background-color: #f1f5f9 !important;
    font-weight: 700;
    font-size: 13px;
    color: #0f172a;
    border-top: 2px solid #cbd5e1;
  }}
  .grand-total-row td {{
    padding: 10px 8px;
  }}
</style>
</head>
<body>
  <div id="div-table-1">
    {t1_html}
  </div>
  <div id="div-table-2">
    {t2_html}
  </div>
  <div id="div-table-3">
    {t3_html}
  </div>
  <div id="div-table-4">
    {t4_html}
  </div>
</body>
</html>
"""
    html_path = "temp_opr_tables.html"
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(full_html)

    # 4. Take Element Screenshots using Playwright
    print("-> Đang chụp ảnh các bảng đẹp bằng Playwright...")
    img_paths = {
        "table1": "table1_opr_AM.png",
        "table2": "table2_opr_9h19h.png",
        "table3": "table3_opr_19h9h.png",
        "table4": "table4_tytrong_opr.png"
    }

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch()
            page = browser.new_page()
            page.set_viewport_size({"width": 900, "height": 2200})
            page.goto(f"file:///{os.path.abspath(html_path)}")
            page.wait_for_timeout(400)
            
            page.locator("#div-table-1 .table-card").screenshot(path=img_paths["table1"])
            page.locator("#div-table-2 .table-card").screenshot(path=img_paths["table2"])
            page.locator("#div-table-3 .table-card").screenshot(path=img_paths["table3"])
            page.locator("#div-table-4 .table-card").screenshot(path=img_paths["table4"])
            
            browser.close()
        print("✔️ Đã chụp thành công 4 ảnh bảng báo cáo visual mới siêu đẹp.")
    except Exception as e:
        print(f"❌ Lỗi chụp ảnh Playwright: {e}")
        sys.exit(1)

    # 5. Broadcast to Telegram & GTalk
    print("📡 Đang gửi báo cáo qua Telegram và GTalk...")

    diff1 = r1['totalOpr'] - r1['prevTotalOpr']
    diff1_str = f"🟢 +{diff1*100:.2f}%" if diff1 > 0 else (f"🔴 {diff1*100:.2f}%" if diff1 < 0 else "0.00%")

    diff2 = r2['totalOpr'] - r2['prevTotalOpr']
    diff2_str = f"🟢 +{diff2*100:.2f}%" if diff2 > 0 else (f"🔴 {diff2*100:.2f}%" if diff2 < 0 else "0.00%")

    diff3 = r3['totalOpr'] - r3['prevTotalOpr']
    diff3_str = f"🟢 +{diff3*100:.2f}%" if diff3 > 0 else (f"🔴 {diff3*100:.2f}%" if diff3 < 0 else "0.00%")

    # Bảng 1: %OPR TTS
    msg1_tele = f"📊 <b>BẢNG 1: %OPR TTS THEO AM - KPI 80%</b>\n" \
                f"⏱️ <b>Mốc:</b> {time_label} ngày {date_label}\n" \
                f"📦 <b>Tổng đơn ltc:</b> {int(r1['totalVol']):,} đơn\n" \
                f"🟢 <b>Đúng hạn:</b> {int(r1['totalOt']):,} đơn\n" \
                f"❌ <b>Đơn trễ:</b> {int(r1['totalLate']):,} đơn\n" \
                f"📈 <b>%OPR TTS ({date_n_label}):</b> {r1['totalOpr'] * 100:.2f}%\n" \
                f"🔄 <b>So với {date_n1_label} ({r1['prevTotalOpr'] * 100:.2f}%):</b> {diff1_str}"

    msg1_gtalk = msg1_tele + "\n\nChi tiết đơn lỗi theo AM (<a href=\"https://docs.google.com/spreadsheets/d/1d3Yeu-5mBE8w5i89_dyJ0ICl1GNP7WgZrH1oQfc5j0s/edit?gid=0#gid=0\"><b>xem chi tiết</b></a>)"

    # Bảng 2: %OPR TTS
    msg2 = f"📊 <b>BẢNG 2: %OPR TTS TẠO TỪ 9H-19H</b>\n" \
           f"⏱️ <b>Mốc:</b> {time_label} ngày {date_label}\n" \
           f"📦 <b>Tổng đơn ltc:</b> {int(r2['totalVol']):,} đơn\n" \
           f"🟢 <b>Đúng hạn:</b> {int(r2['totalOt']):,} đơn\n" \
           f"❌ <b>Đơn trễ:</b> {int(r2['totalLate']):,} đơn\n" \
           f"📈 <b>%OPR TTS ({date_n_label}):</b> {r2['totalOpr'] * 100:.2f}%\n" \
           f"🔄 <b>So với {date_n1_label} ({r2['prevTotalOpr'] * 100:.2f}%):</b> {diff2_str}"

    # Bảng 3: %OPR TTS
    msg3 = f"📊 <b>BẢNG 3: %OPR TTS TẠO TỪ 19H-9H (Ca đêm)</b>\n" \
           f"⏱️ <b>Mốc:</b> {time_label} ngày {date_label}\n" \
           f"📦 <b>Tổng đơn ltc:</b> {int(r3['totalVol']):,} đơn\n" \
           f"🟢 <b>Đúng hạn:</b> {int(r3['totalOt']):,} đơn\n" \
           f"❌ <b>Đơn trễ:</b> {int(r3['totalLate']):,} đơn\n" \
           f"📈 <b>%OPR TTS ca đêm ({date_n_label}):</b> {r3['totalOpr'] * 100:.2f}%\n" \
           f"🔄 <b>So với {date_n1_label} ({r3['prevTotalOpr'] * 100:.2f}%):</b> {diff3_str}"

    # Bảng 4: %OPR TTS
    pct9 = r4["totalLate9to19"] / r4["totalLateAll"] * 100 if r4["totalLateAll"] > 0 else 0.0
    pct19 = r4["totalLate19to9"] / r4["totalLateAll"] * 100 if r4["totalLateAll"] > 0 else 0.0
    msg4 = f"📊 <b>BẢNG 4: TỶ TRỌNG LỖI OPR TTS THEO KHUNG GIỜ TẠO</b>\n" \
           f"⏱️ <b>Mốc:</b> {time_label} ngày {date_label}\n" \
           f"❌ <b>Tổng đơn trễ ({date_n_label}):</b> {int(r4['totalLateAll']):,} đơn\n" \
           f"  • Ca ngày (9h-19h): <b>{int(r4['totalLate9to19']):,} đơn</b> ({pct9:.2f}%)\n" \
           f"  • Ca đêm (19h-9h): <b>{int(r4['totalLate19to9']):,} đơn</b> ({pct19:.2f}%)"

    # Send Table 1
    send_photo_telegram(img_paths["table1"], msg1_tele)
    send_photo_gtalk(img_paths["table1"], msg1_gtalk)
    time.sleep(2)

    # Send Table 2
    send_photo_telegram(img_paths["table2"], msg2)
    send_photo_gtalk(img_paths["table2"], msg2)
    time.sleep(2)

    # Send Table 3
    send_photo_telegram(img_paths["table3"], msg3)
    send_photo_gtalk(img_paths["table3"], msg3)
    time.sleep(2)

    # Send Table 4
    send_photo_telegram(img_paths["table4"], msg4)
    send_photo_gtalk(img_paths["table4"], msg4)

    # Clean up local temporary files
    try:
        os.remove(html_path)
        for p in img_paths.values():
            if os.path.exists(p):
                os.remove(p)
    except Exception:
        pass

    print("✔️ Đã gửi tất cả báo cáo sang Telegram và GTalk thành công!")

if __name__ == "__main__":
    main()
