# -*- coding: utf-8 -*-
"""
Script: report_tongquan.py
Processes NTB Overview and Top Post Offices performance data from Google Sheets,
renders beautiful visual tables using Playwright, and broadcasts them to Telegram & GTalk.
"""

import os
import sys
import time
import urllib3
import requests
import unicodedata
import pandas as pd
import numpy as np
from datetime import datetime
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

# ============ CONFIG & CONSTANTS ============
SPREADSHEET_ID = "1DAwY-46twFrHIs77R4p4IMuIZ6JTE-e58Aj-9Kcr5Jk"
TELEGRAM_TOKEN = "8570130113:AAGXRiUaKBknVpgtm1_i9ZA47JRjAXmB21M"
TELEGRAM_CHAT_ID = "-5058464865"
GTALK_OA_TOKEN = "2067164759710552066:RfgqBJY4QtV18udu1wMEfhmoRNI4hgBv"
GTALK_CHANNEL_ID = "2067164759710552066"  # Channel ID used by other scripts

SERVICE_ACCOUNT_CANDIDATES = [
    r"C:\Users\lap4all\Desktop\Backlog_Automation\credentials.json",
    r"C:\Users\lap4all\Downloads\credentials.json",
    r"C:\Users\lap4all\Downloads\service_account.json",
    r"C:\Users\lap4all\Desktop\credentials.json",
    "credentials.json",
    "service_account.json",
]

# Load dotenv from desktop configuration if available
env_path = r"c:\Users\lap4all\Desktop\New folder\.env"
if os.path.exists(env_path):
    load_dotenv(dotenv_path=env_path, override=True)
else:
    load_dotenv()

# Override config if defined in env
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN") or os.environ.get("TELEGRAM_BOT_TOKEN") or TELEGRAM_TOKEN
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID") or TELEGRAM_CHAT_ID
GTALK_OA_TOKEN = os.environ.get("GTALK_OA_TOKEN") or GTALK_OA_TOKEN
GTALK_CHANNEL_ID = os.environ.get("GTALK_CHANNEL_ID") or GTALK_CHANNEL_ID

def find_service_account_file():
    for p in SERVICE_ACCOUNT_CANDIDATES:
        if os.path.isfile(p):
            return p
    raise FileNotFoundError("Không tìm thấy file credentials.json.")

# ============ FORMATTING HELPERS ============
def format_vn_int(val):
    if val is None or pd.isna(val) or val == "":
        return ""
    try:
        s = f"{int(round(float(str(val).replace('.', '').replace(',', '.')))):,}"
        return s.replace(",", ".")
    except Exception:
        return str(val)

def format_vn_pct(val):
    if val is None or pd.isna(val) or val == "":
        return ""
    try:
        # Check if already formatted as a percentage string
        if isinstance(val, str) and "%" in val:
            return val.strip()
        # Handle Vietnamese decimal comma
        clean_val = str(val).replace(',', '.')
        f_val = float(clean_val)
        # If the number is like 84.18 (out of 100) or 0.8418
        if f_val > 1.0:
            s = f"{f_val:.2f}"
        else:
            s = f"{f_val * 100:.2f}"
        return s.replace(".", ",") + "%"
    except Exception:
        return str(val)

def parse_pct_float(val):
    if val is None or val == "":
        return 0.0
    try:
        if isinstance(val, str):
            val_clean = val.replace("%", "").replace(",", ".").strip()
            num = float(val_clean)
            if num <= 1.0:
                return num * 100.0
            return num
        else:
            num = float(val)
            if num <= 1.0:
                return num * 100.0
            return num
    except:
        return 0.0

# ============ SHEET PARSERS ============
def parse_tongquan_sheet(rows):
    ntb_headers = None
    ntb_data = []
    ntb_total = None
    
    am_headers = None
    am_data = []
    am_total = None
    
    am_header_row_idx = None
    am_total_row_idx = None
    
    mode = "NTB_SEARCH"
    
    for r_idx, r in enumerate(rows):
        if not r or all(str(x).strip() == "" for x in r):
            continue
        
        # Clean row
        r_clean = [str(x).strip() for x in r]
        first_cell = r_clean[0]
        
        if first_cell == "Vùng":
            ntb_headers = r_clean
            mode = "NTB_DATA"
            continue
        elif first_cell == "Tổng cộng" and mode == "NTB_DATA":
            ntb_total = r_clean
            mode = "AM_SEARCH"
            continue
        elif first_cell == "AM" and am_headers is None:
            am_headers = r_clean
            am_header_row_idx = r_idx + 1
            mode = "AM_DATA"
            continue
        elif first_cell == "Tổng cộng" and mode == "AM_DATA":
            am_total = r_clean
            am_total_row_idx = r_idx + 1
            mode = "END"
            break
            
        if mode == "NTB_DATA":
            ntb_data.append(r_clean)
        elif mode == "AM_DATA":
            am_data.append(r_clean)
            
    return {
        "ntb_headers": ntb_headers,
        "ntb_data": ntb_data,
        "ntb_total": ntb_total,
        "am_headers": am_headers,
        "am_data": am_data,
        "am_total": am_total,
        "am_header_row_idx": am_header_row_idx,
        "am_total_row_idx": am_total_row_idx
    }


def aggregate_am_stats(df_date, am_list):
    am_data = []
    
    for am in am_list:
        df_am = df_date[df_date['AM'] == am]
        
        vol = int(df_am['Volume'].sum())
        hm = int(df_am['Hàng Mới Về Trong Ngày'].sum())
        gan = int(df_am['Sản Lượng Gán'].sum())
        chua_gan = int(df_am['Sản Lượng Chưa Gán'].sum())
        gtc = int(df_am['Sản Lượng Giao Thành Công'].sum())
        tra = int(df_am['Sản Lượng Chuyển Trả'].sum())
        ton = int(df_am['Sản Lượng Tồn'].sum())
        
        pct_gan = gan / vol if vol > 0 else 0.0
        pct_chua_gan = chua_gan / vol if vol > 0 else 0.0
        pct_gtc = gtc / vol if vol > 0 else 0.0
        pct_tra = tra / vol if vol > 0 else 0.0
        pct_ton = ton / vol if vol > 0 else 0.0
        
        time_str = df_date['Time'].iloc[0] if not df_date.empty else ""
        
        am_data.append([
            am, time_str, vol, hm, gan, chua_gan, gtc, tra, ton,
            pct_gan, pct_chua_gan, pct_gtc, pct_tra, pct_ton
        ])
        
    total_vol = sum(int(row[2]) for row in am_data)
    total_hm = sum(int(row[3]) for row in am_data)
    total_gan = sum(int(row[4]) for row in am_data)
    total_chua_gan = sum(int(row[5]) for row in am_data)
    total_gtc = sum(int(row[6]) for row in am_data)
    total_tra = sum(int(row[7]) for row in am_data)
    total_ton = sum(int(row[8]) for row in am_data)
    
    total_pct_gan = total_gan / total_vol if total_vol > 0 else 0.0
    total_pct_chua_gan = total_chua_gan / total_vol if total_vol > 0 else 0.0
    total_pct_gtc = total_gtc / total_vol if total_vol > 0 else 0.0
    total_pct_tra = total_tra / total_vol if total_vol > 0 else 0.0
    total_pct_ton = total_ton / total_vol if total_vol > 0 else 0.0
    
    am_total = [
        "Tổng cộng", "", total_vol, total_hm, total_gan, total_chua_gan, total_gtc, total_tra, total_ton,
        total_pct_gan, total_pct_chua_gan, total_pct_gtc, total_pct_tra, total_pct_ton
    ]
    
    return am_data, am_total

    
def update_all_pivot_tables(client, sh, unique_data_dates, df):
    if len(unique_data_dates) < 2:
        print("⚠️ Không đủ ngày trong dữ liệu để cập nhật pivot tables.")
        return
        
    latest_8_dates = unique_data_dates[-8:]
    latest_date = unique_data_dates[-1]
    prev_date = unique_data_dates[-2]
    
    # Check if latest_date has Ca 2 and Tồn data in df
    df_latest = df[df['Time'] == latest_date]
    
    has_latest_ton = False
    if 'Loại Hàng' in df_latest.columns:
        has_latest_ton = df_latest[df_latest['Loại Hàng'] == 'Hàng Tồn']['Volume'].sum() > 0
    else:
        has_latest_ton = df_latest['Sản Lượng Tồn'].sum() > 0
        
    has_latest_ca2 = False
    if 'Loại Hàng' in df_latest.columns:
        has_latest_ca2 = df_latest[df_latest['Loại Hàng'] == 'Hàng Mới Ca 2']['Volume'].sum() > 0
        
    ton_date = latest_date if has_latest_ton else prev_date
    ca2_date = latest_date if has_latest_ca2 else prev_date
    
    print(f"📊 Đang tự động cập nhật các Pivot Tables trên Sheets...")
    print(f"  • Mốc ngày mới nhất: {latest_date}")
    print(f"  • Mốc ngày hôm trước: {prev_date}")
    print(f"  • Chọn ngày cho pivot Tồn: {ton_date} (có data: {has_latest_ton})")
    print(f"  • Chọn ngày cho pivot Ca 2: {ca2_date} (có data: {has_latest_ca2})")
    
    try:
        ws_tq = sh.worksheet("Tổng quan")
        ws_ca = sh.worksheet("Ca1 - Ca2 - Tồn")
        
        # 1. Fetch cell metadata for the pivot tables
        ranges = [
            f"'{ws_tq.title}'!A1:A1",
            f"'{ws_tq.title}'!A12:A12",
            f"'{ws_tq.title}'!A60:A60",
            f"'{ws_ca.title}'!A2:A2",
            f"'{ws_ca.title}'!Q2:Q2",
            f"'{ws_ca.title}'!AG2:AG2",
            f"'{ws_ca.title}'!AW2:AW2"
        ]
        
        metadata = client.http_client.spreadsheets_get(
            sh.id, 
            params={"includeGridData": True, "ranges": ranges}
        )
        
        sheets = metadata.get("sheets", [])
        if not sheets:
            print("⚠️ Không lấy được metadata của các pivot tables.")
            return
            
        # Map sheetId to title
        sheet_map = {s["properties"]["title"]: s["properties"]["sheetId"] for s in sheets}
        sheet_id_tq = sheet_map.get("Tổng quan")
        sheet_id_ca = sheet_map.get("Ca1 - Ca2 - Tồn")
        
        # Extract rowData from both sheets in the metadata
        cell_map = {}
        for sheet_data in sheets:
            title = sheet_data["properties"]["title"]
            sheet_id = sheet_data["properties"]["sheetId"]
            data_ranges = sheet_data.get("data", [])
            for dr in data_ranges:
                start_row = dr.get("startRow", 0)
                start_col = dr.get("startColumn", 0)
                row_data = dr.get("rowData", [])
                if row_data:
                    values = row_data[0].get("values", [])
                    if values:
                        cell_map[(sheet_id, start_row, start_col)] = values[0]
                        
        update_requests = []
        
        # Define updates config
        configs = [
            {"sheet_id": sheet_id_tq, "coord": (0, 0), "label": "Tổng quan (A1)", "dates": latest_8_dates},
            {"sheet_id": sheet_id_tq, "coord": (11, 0), "label": "Tổng quan AM (A12)", "dates": [latest_date]},
            {"sheet_id": sheet_id_tq, "coord": (59, 0), "label": "Tổng quan Tỉnh (A60)", "dates": [latest_date]},
            {"sheet_id": sheet_id_ca, "coord": (1, 0), "label": "Ca1 + Tồn (A2)", "dates": [latest_date]},
            {"sheet_id": sheet_id_ca, "coord": (1, 16), "label": "Tồn (Q2)", "dates": [ton_date]},
            {"sheet_id": sheet_id_ca, "coord": (1, 32), "label": "Ca1 (AG2)", "dates": [latest_date]},
            {"sheet_id": sheet_id_ca, "coord": (1, 48), "label": "Ca2 (AW2)", "dates": [ca2_date]}
        ]
        
        for cfg in configs:
            sid = cfg["sheet_id"]
            coord = cfg["coord"]
            label = cfg["label"]
            target_dates = cfg["dates"]
            
            cell = cell_map.get((sid, coord[0], coord[1]))
            if not cell or "pivotTable" not in cell:
                print(f"⚠️ Không tìm thấy pivot table cho {label}")
                continue
                
            pivot = cell["pivotTable"]
            
            # Update criteria["3"]
            if "criteria" in pivot and "3" in pivot["criteria"]:
                pivot["criteria"]["3"]["visibleValues"] = target_dates
            elif "criteria" in pivot and 3 in pivot["criteria"]:
                pivot["criteria"][3]["visibleValues"] = target_dates
                
            # Update filterSpecs
            if "filterSpecs" in pivot:
                for spec in pivot["filterSpecs"]:
                    if spec.get("columnOffsetIndex") == 3:
                        if "filterCriteria" in spec:
                            spec["filterCriteria"]["visibleValues"] = target_dates
                            
            update_requests.append({
                "updateCells": {
                    "rows": [{"values": [{"pivotTable": pivot}]}],
                    "fields": "pivotTable",
                    "range": {
                        "sheetId": sid,
                        "startRowIndex": coord[0],
                        "endRowIndex": coord[0] + 1,
                        "startColumnIndex": coord[1],
                        "endColumnIndex": coord[1] + 1
                    }
                }
            })
            
        if update_requests:
            # Clear cells on 'Tổng quan' to allow pivot table expansion (including AM and Tỉnh tables)
            try:
                ws_tq.batch_clear(["A2:A10", "B2:N10", "A13:A40", "B13:N40", "A61:A75", "B61:N75"])
            except Exception as e:
                print(f"⚠️ Cảnh báo khi xóa vùng trống: {e}")
                
            # Send batchUpdate requests
            client.http_client.batch_update(sh.id, {"requests": update_requests})
            print("✔️ Đã cập nhật xong tất cả các pivot tables.")
        else:
            print("⚠️ Không có yêu cầu cập nhật pivot table nào được tạo.")
            
    except Exception as e:
        print(f"❌ Lỗi khi cập nhật pivot tables: {e}")

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
        print(f"✅ Đã gửi ảnh {os.path.basename(image_path)} lên Telegram thành công.")
    else:
        print(f"❌ Lỗi gửi Telegram: {result}")
    return result

# ============ GTALK SENDER ============
def send_photo_gtalk(image_path, caption=""):
    if not GTALK_OA_TOKEN or not GTALK_CHANNEL_ID:
        print("⚠️ Không tìm thấy GTALK_OA_TOKEN hoặc GTALK_CHANNEL_ID. Bỏ qua gửi GTalk.")
        return False

    print(f"📡 Đang gửi ảnh {os.path.basename(image_path)} sang GTalk...")
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
                print(f"✅ Đã gửi ảnh {os.path.basename(image_path)} sang GTalk thành công!")
                return True
            else:
                print(f"❌ Lỗi gửi tin nhắn GTalk API: {res_data_send.get('error')}")
        else:
            print(f"❌ Lỗi HTTP {res_send.status_code}: {res_send.text}")
    except Exception as e:
        print(f"❌ Lỗi kết nối khi gửi tin nhắn GTalk: {e}")
    return False

# ============ HTML BUILDERS ============
def build_overview_html(data, time_label, date_label):
    # NTB Data rows
    ntb_rows_html = ""
    for idx, r in enumerate(data["ntb_data"]):
        bg = "#ffffff" if idx % 2 == 0 else "#f8fafc"
        
        # Color coding for %GTC (index 11)
        try:
            gtc_num = parse_pct_float(r[11])
            if gtc_num >= 63.0:
                gtc_class = "gtc-good"
            elif gtc_num >= 59.0:
                gtc_class = "gtc-warning"
            else:
                gtc_class = "gtc-bad"
        except:
            gtc_class = ""
            
        ntb_rows_html += f"""
        <tr style="background-color: {bg};">
            <td style="font-weight: 600; text-align: left; color: #0f172a;">{r[0]}</td>
            <td style="color: #475569;">{r[1]}</td>
            <td>{format_vn_int(r[2])}</td>
            <td>{format_vn_int(r[3])}</td>
            <td>{format_vn_int(r[4])}</td>
            <td>{format_vn_int(r[5])}</td>
            <td>{format_vn_int(r[6])}</td>
            <td>{format_vn_int(r[7])}</td>
            <td style="font-weight: 500;">{format_vn_int(r[8])}</td>
            <td>{format_vn_pct(r[9])}</td>
            <td>{format_vn_pct(r[10])}</td>
            <td class="{gtc_class}">{format_vn_pct(r[11])}</td>
            <td>{format_vn_pct(r[12])}</td>
            <td>{format_vn_pct(r[13])}</td>
        </tr>
        """
        
    # NTB Total
    t = data["ntb_total"]
    ntb_total_html = f"""
    <tr class="total-row">
        <td style="text-align: left;">{t[0]}</td>
        <td></td>
        <td>{format_vn_int(t[2])}</td>
        <td>{format_vn_int(t[3])}</td>
        <td>{format_vn_int(t[4])}</td>
        <td>{format_vn_int(t[5])}</td>
        <td>{format_vn_int(t[6])}</td>
        <td>{format_vn_int(t[7])}</td>
        <td>{format_vn_int(t[8])}</td>
        <td>{format_vn_pct(t[9])}</td>
        <td>{format_vn_pct(t[10])}</td>
        <td>{format_vn_pct(t[11])}</td>
        <td>{format_vn_pct(t[12])}</td>
        <td>{format_vn_pct(t[13])}</td>
    </tr>
    """

    # AM Data rows
    am_rows_html = ""
    for idx, r in enumerate(data["am_data"]):
        bg = "#ffffff" if idx % 2 == 0 else "#f8fafc"
        
        # Color coding for %GTC (index 11)
        try:
            gtc_num = parse_pct_float(r[11])
            if gtc_num >= 65.0:
                gtc_class = "gtc-good"
            elif gtc_num >= 60.0:
                gtc_class = "gtc-warning"
            else:
                gtc_class = "gtc-bad"
        except:
            gtc_class = ""
            
        am_rows_html += f"""
        <tr style="background-color: {bg};">
            <td style="font-weight: 600; text-align: left; color: #0f172a;">{r[0]}</td>
            <td style="color: #64748b;">{r[1]}</td>
            <td>{format_vn_int(r[2])}</td>
            <td>{format_vn_int(r[3])}</td>
            <td>{format_vn_int(r[4])}</td>
            <td>{format_vn_int(r[5])}</td>
            <td>{format_vn_int(r[6])}</td>
            <td>{format_vn_int(r[7])}</td>
            <td style="font-weight: 500;">{format_vn_int(r[8])}</td>
            <td>{format_vn_pct(r[9])}</td>
            <td>{format_vn_pct(r[10])}</td>
            <td class="{gtc_class}">{format_vn_pct(r[11])}</td>
            <td>{format_vn_pct(r[12])}</td>
            <td>{format_vn_pct(r[13])}</td>
        </tr>
        """
        
    # AM Total
    at = data["am_total"]
    am_total_html = f"""
    <tr class="total-row">
        <td style="text-align: left;">{at[0]}</td>
        <td></td>
        <td>{format_vn_int(at[2])}</td>
        <td>{format_vn_int(at[3])}</td>
        <td>{format_vn_int(at[4])}</td>
        <td>{format_vn_int(at[5])}</td>
        <td>{format_vn_int(at[6])}</td>
        <td>{format_vn_int(at[7])}</td>
        <td>{format_vn_int(at[8])}</td>
        <td>{format_vn_pct(at[9])}</td>
        <td>{format_vn_pct(at[10])}</td>
        <td>{format_vn_pct(at[11])}</td>
        <td>{format_vn_pct(at[12])}</td>
        <td>{format_vn_pct(at[13])}</td>
    </tr>
    """

    html = f"""<!DOCTYPE html>
    <html>
    <head>
    <meta charset="utf-8">
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <style>
        body {{
            font-family: 'Inter', sans-serif;
            background-color: #f1f5f9;
            margin: 0;
            padding: 24px;
            display: flex;
            justify-content: center;
        }}
        .container {{
            width: 1400px;
            background-color: #ffffff;
            border-radius: 16px;
            box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.05), 0 4px 6px -4px rgba(0, 0, 0, 0.05);
            padding: 28px;
            border: 1px solid #e2e8f0;
        }}
        .header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 2px dashed #e2e8f0;
            padding-bottom: 20px;
            margin-bottom: 24px;
        }}
        .header-title h1 {{
            font-size: 24px;
            font-weight: 700;
            color: #0f172a;
            margin: 0;
            background: linear-gradient(135deg, #1e3a8a, #3b82f6);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }}
        .header-title p {{
            font-size: 13px;
            color: #64748b;
            margin: 4px 0 0 0;
            font-weight: 500;
        }}
        .badge-update {{
            background-color: #eff6ff;
            border: 1px solid #bfdbfe;
            color: #1d4ed8;
            padding: 6px 14px;
            border-radius: 9999px;
            font-size: 13px;
            font-weight: 600;
        }}
        .section-title {{
            font-size: 16px;
            font-weight: 700;
            color: #1e293b;
            margin-top: 10px;
            margin-bottom: 14px;
            display: flex;
            align-items: center;
            gap: 8px;
        }}
        .section-title::before {{
            content: '';
            display: inline-block;
            width: 4px;
            height: 18px;
            background-color: #2563eb;
            border-radius: 2px;
        }}
        table {{
            width: 100%;
            border-collapse: separate;
            border-spacing: 0;
            border: 1px solid #e2e8f0;
            border-radius: 8px;
            overflow: hidden;
            font-size: 12.5px;
            margin-bottom: 32px;
        }}
        th {{
            background: linear-gradient(180deg, #1e293b, #0f172a);
            color: #ffffff;
            font-weight: 600;
            text-align: center;
            padding: 12px 10px;
            border-bottom: 1px solid #334155;
            font-size: 12px;
        }}
        th:first-child {{
            text-align: left;
            padding-left: 16px;
        }}
        td {{
            padding: 10px 10px;
            text-align: center;
            border-bottom: 1px solid #e2e8f0;
            color: #0f172a;
            font-weight: 600;
        }}
        td:first-child {{
            padding-left: 16px;
        }}
        tr:last-child td {{
            border-bottom: none;
        }}
        .total-row {{
            background-color: #f1f5f9 !important;
            font-weight: 700;
        }}
        .total-row td {{
            color: #0f172a;
            border-top: 2px solid #cbd5e1;
        }}
        /* Color statuses */
        .gtc-good {{
            background-color: #dcfce7 !important;
            color: #15803d !important;
            font-weight: 700 !important;
        }}
        .gtc-warning {{
            background-color: #fef9c3 !important;
            color: #a16207 !important;
            font-weight: 700 !important;
        }}
        .gtc-bad {{
            background-color: #fee2e2 !important;
            color: #b91c1c !important;
            font-weight: 700 !important;
        }}
    </style>
    </head>
    <body>
        <div id="overview-container" class="container">
            <div class="header">
                <div class="header-title">
                    <h1>BÁO CÁO VẬN HÀNH TỔNG QUAN</h1>
                </div>
                <div class="badge-update">
                    Mốc cập nhật: {time_label} ngày {date_label}
                </div>
            </div>
            
            <div class="section-title">HIỆU SUẤT VẬN HÀNH TOÀN VÙNG (NTB)</div>
            <table>
                <thead>
                    <tr>
                        <th style="width: 80px;">Vùng</th>
                        <th style="width: 140px;">Time</th>
                        <th>Volume</th>
                        <th>Hàng mới về</th>
                        <th>Gán</th>
                        <th>Chưa Gán</th>
                        <th>GTC</th>
                        <th>Chuyển Trả</th>
                        <th>Tồn</th>
                        <th>% Gán</th>
                        <th>% Chưa Gán</th>
                        <th>% GTC</th>
                        <th>% Chuyển trả</th>
                        <th>% Tồn</th>
                    </tr>
                </thead>
                <tbody>
                    {ntb_rows_html}
                    {ntb_total_html}
                </tbody>
            </table>
            
            <div class="section-title">CHI TIẾT HIỆU SUẤT THEO AM (Mốc {data["am_data"][0][1] if len(data["am_data"]) > 0 else ""})</div>
            <table>
                <thead>
                    <tr>
                        <th style="width: 180px;">AM</th>
                        <th style="width: 140px;">Time</th>
                        <th>Volume</th>
                        <th>Hàng mới về</th>
                        <th>Gán</th>
                        <th>Chưa Gán</th>
                        <th>GTC</th>
                        <th>Chuyển Trả</th>
                        <th>Tồn</th>
                        <th>% Gán</th>
                        <th>% Chưa Gán</th>
                        <th>% GTC</th>
                        <th>% Chuyển trả</th>
                        <th>% Tồn</th>
                    </tr>
                </thead>
                <tbody>
                    {am_rows_html}
                    {am_total_html}
                </tbody>
            </table>
        </div>
    </body>
    </html>
    """
    return html

def build_top_offices_html(title, df, is_best, date_str):
    rows_html = ""
    for idx, r in enumerate(df.to_dict(orient='records')):
        bg = "#ffffff" if idx % 2 == 0 else "#f8fafc"
        
        # Color coding for %GTC
        gtc_pct = r['%GTC'] * 100
        if is_best:
            # Scale greens
            if gtc_pct >= 85:
                gtc_style = "background-color: #dcfce7; color: #15803d; font-weight: 600;"
            elif gtc_pct >= 75:
                gtc_style = "background-color: #f0fdf4; color: #166534; font-weight: 600;"
            else:
                gtc_style = "background-color: #fef9c3; color: #854d0e; font-weight: 600;"
        else:
            # Scale reds
            if gtc_pct < 30:
                gtc_style = "background-color: #fee2e2; color: #991b1b; font-weight: 700;"
            elif gtc_pct < 45:
                gtc_style = "background-color: #ffedd5; color: #c2410c; font-weight: 600;"
            else:
                gtc_style = "background-color: #fef9c3; color: #854d0e; font-weight: 600;"
                
        rows_html += f"""
        <tr style="background-color: {bg};">
            <td style="font-weight: 600; color: #1e293b; text-align: center;">{idx+1:02d}</td>
            <td style="text-align: left; font-weight: 500; color: #0f172a;">{r['Chi tiết']}</td>
            <td style="text-align: left; color: #475569; font-weight: 500;">{r.get('AM', '')}</td>
            <td style="color: #64748b;">{r['Time']}</td>
            <td>{format_vn_int(r['Volume'])}</td>
            <td>{format_vn_int(r['Hàng Mới Về Trong Ngày'])}</td>
            <td>{format_vn_int(r['Sản Lượng Gán'])}</td>
            <td>{format_vn_int(r['Sản Lượng Giao Thành Công'])}</td>
            <td>{format_vn_int(r['Sản Lượng Tồn'])}</td>
            <td>{format_vn_pct(r['%Gán'])}</td>
            <td style="{gtc_style}">{format_vn_pct(r['%GTC'])}</td>
        </tr>
        """
        
    theme_gradient = "linear-gradient(135deg, #15803d, #22c55e)" if is_best else "linear-gradient(135deg, #991b1b, #ef4444)"
    badge_bg = "#dcfce7" if is_best else "#fee2e2"
    badge_color = "#15803d" if is_best else "#991b1b"
    
    html = f"""<!DOCTYPE html>
    <html>
    <head>
    <meta charset="utf-8">
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <style>
        body {{
            font-family: 'Inter', sans-serif;
            background-color: #f1f5f9;
            margin: 0;
            padding: 24px;
            display: flex;
            justify-content: center;
        }}
        .container {{
            width: 1100px;
            background-color: #ffffff;
            border-radius: 16px;
            box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.05), 0 4px 6px -4px rgba(0, 0, 0, 0.05);
            padding: 28px;
            border: 1px solid #e2e8f0;
        }}
        .header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 2px dashed #e2e8f0;
            padding-bottom: 20px;
            margin-bottom: 24px;
        }}
        .header-title h1 {{
            font-size: 22px;
            font-weight: 700;
            color: #0f172a;
            margin: 0;
            background: {theme_gradient};
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }}
        .header-title p {{
            font-size: 13px;
            color: #64748b;
            margin: 4px 0 0 0;
            font-weight: 500;
        }}
        .badge-type {{
            background-color: {badge_bg};
            color: {badge_color};
            padding: 6px 16px;
            border-radius: 9999px;
            font-size: 13px;
            font-weight: 700;
            text-transform: uppercase;
            border: 1px solid rgba(0, 0, 0, 0.05);
        }}
        table {{
            width: 100%;
            border-collapse: separate;
            border-spacing: 0;
            border: 1px solid #e2e8f0;
            border-radius: 8px;
            overflow: hidden;
            font-size: 12.5px;
        }}
        th {{
            background: linear-gradient(180deg, #1e293b, #0f172a);
            color: #ffffff;
            font-weight: 600;
            text-align: center;
            padding: 12px 8px;
            border-bottom: 1px solid #334155;
            font-size: 12px;
        }}
        th:first-child {{
            border-top-left-radius: 8px;
        }}
        th:last-child {{
            border-top-right-radius: 8px;
        }}
        td {{
            padding: 10px 8px;
            text-align: center;
            border-bottom: 1px solid #e2e8f0;
            color: #0f172a;
            font-weight: 600;
        }}
        tr:last-child td {{
            border-bottom: none;
        }}
    </style>
    </head>
    <body>
        <div id="capture-container" class="container">
            <div class="header">
                <div class="header-title">
                    <h1>{title}</h1>
                    <p>Bưu cục có sản lượng lớn &geq; 300 đơn hàng</p>
                </div>
                <div class="badge-type">
                    Ngày {date_str}
                </div>
            </div>
            <table>
                <thead>
                    <tr>
                        <th style="width: 40px;">Hạng</th>
                        <th style="text-align: left; width: 220px;">Bưu cục</th>
                        <th style="text-align: left; width: 120px;">AM</th>
                        <th style="width: 120px;">Time</th>
                        <th>Volume</th>
                        <th>Hàng mới về</th>
                        <th>Gán</th>
                        <th>GTC</th>
                        <th>Tồn</th>
                        <th>% Gán</th>
                        <th style="width: 90px;">% GTC</th>
                    </tr>
                </thead>
                <tbody>
                    {rows_html}
                </tbody>
            </table>
        </div>
    </body>
    </html>
    """
    return html

# ============ MAIN PIPELINE ============
def main():
    print("Đang đọc credentials...")
    try:
        json_file = find_service_account_file()
        print(f"✔️ Đã tìm thấy credentials tại: {json_file}")
    except Exception as e:
        print(f"❌ Lỗi: {e}")
        sys.exit(1)

    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_file(json_file, scopes=scope)
    client = gspread.authorize(creds)

    print(f"Đang kết nối tới spreadsheet ID: {SPREADSHEET_ID}...")
    try:
        sh = client.open_by_key(SPREADSHEET_ID)
        print(f"✔️ Đã kết nối thành công tới: '{sh.title}'")
    except Exception as e:
        print(f"❌ Lỗi mở spreadsheet: {e}")
        sys.exit(1)

    # 1. Load "Tổng quan" sheet reference
    print("📖 Kết nối tab 'Tổng quan'...")
    try:
        ws_tq = sh.worksheet("Tổng quan")
    except Exception as e:
        print(f"❌ Lỗi kết nối tab 'Tổng quan': {e}")
        sys.exit(1)

    # 2. Load "Data" sheet and convert to DataFrame
    print("📖 Đọc dữ liệu từ tab 'Data'...")
    try:
        ws_data = sh.worksheet("Data")
    except Exception as e:
        print(f"❌ Lỗi kết nối tab 'Data': {e}")
        sys.exit(1)

    max_retries = 10
    retry_delay = 5
    data_rows = []
    
    for attempt in range(max_retries):
        try:
            data_rows = ws_data.get_all_values()
            df_temp = pd.DataFrame(data_rows[1:], columns=data_rows[0])
            
            unique_dates = df_temp['Time'].unique()
            unique_dates = [d for d in unique_dates if d and str(d).strip()]
            unique_dates.sort(key=lambda x: str(x).split(' ')[0] if x else "")
            
            if not unique_dates:
                print("⚠️ Chưa tìm thấy ngày nào hợp lệ trong cột Time. Sẽ thử lại...")
                time.sleep(retry_delay)
                continue
                
            latest_date = unique_dates[-1]
            df_latest = df_temp[df_temp['Time'] == latest_date]
            
            vol_series = pd.to_numeric(
                df_latest['Volume'].astype(str).str.replace('.', '', regex=False).str.replace(',', '', regex=False),
                errors='coerce'
            ).fillna(0)
            
            gtc_series = pd.to_numeric(
                df_latest['Sản Lượng Giao Thành Công'].astype(str).str.replace('.', '', regex=False).str.replace(',', '', regex=False),
                errors='coerce'
            ).fillna(0)
            
            vol_sum = vol_series.sum()
            gtc_sum = gtc_series.sum()
            
            print(f"🕵️ Kiểm tra dữ liệu ngày {latest_date}: Tổng Volume = {vol_sum:.0f}, Tổng GTC = {gtc_sum:.0f} (Lần thử {attempt + 1}/{max_retries})")
            
            if vol_sum > 0 and gtc_sum == 0:
                print(f"⚠️ Phát hiện Volume > 0 nhưng GTC = 0 (Google Sheets đang tính toán ArrayFormula). Chờ {retry_delay} giây...")
                time.sleep(retry_delay)
            else:
                print("✅ Dữ liệu đã sẵn sàng!")
                break
        except Exception as e:
            print(f"⚠️ Lỗi khi đọc/kiểm tra dữ liệu (Lần thử {attempt + 1}): {e}")
            time.sleep(retry_delay)
    else:
        print("❌ Đã quá thời gian chờ đợi công thức tính toán. Tiếp tục chạy với dữ liệu hiện tại.")
        try:
            data_rows = ws_data.get_all_values()
        except Exception:
            pass

    print(f"✔️ Đã đọc xong {len(data_rows)} dòng dữ liệu thô.")

    print("📊 Đang xử lý data và kiểm tra ngày cập nhật...")
    try:
        df = pd.DataFrame(data_rows[1:], columns=data_rows[0])
        
        # Clean numeric columns in full df first
        numeric_cols = [
            'Volume', 'Sản Lượng Giao Thành Công', 'Sản Lượng Gán', 
            'Sản Lượng Tồn', 'Hàng Mới Về Trong Ngày', 'Sản Lượng Chưa Gán',
            'Sản Lượng Chuyển Trả'
        ]
        for col in numeric_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(
                    df[col].astype(str).str.replace('.', '', regex=False).str.replace(',', '', regex=False), 
                    errors='coerce'
                ).fillna(0).astype(int)
            else:
                df[col] = 0

        # Unique dates in Data sheet
        unique_data_dates = df['Time'].unique()
        unique_data_dates = [d for d in unique_data_dates if d and str(d).strip()]
        unique_data_dates.sort(key=lambda x: str(x).split(' ')[0] if x else "")

        # Update Pivot Tables dynamically
        update_all_pivot_tables(client, sh, unique_data_dates, df)

        # Reload ws_tq values to ensure all rows, indexes and pivot values are fresh
        tq_rows = ws_tq.get_all_values()
        parsed_tq = parse_tongquan_sheet(tq_rows)
        if not parsed_tq["ntb_headers"] or not parsed_tq["ntb_data"]:
            print("❌ Lỗi: Không thể phân tích dữ liệu NTB từ tab 'Tổng quan'.")
            sys.exit(1)



        # Find where "Tổng cộng" row is in ws_tq (dynamically)
        tot_cells = ws_tq.findall("Tổng cộng")
        ntb_tot_cell = None
        for c in tot_cells:
            if c.col == 1:
                ntb_tot_cell = c
                break

        if not ntb_tot_cell:
            print("❌ Không tìm thấy dòng 'Tổng cộng' trong tab 'Tổng quan'.")
            sys.exit(1)

        total_row_idx = ntb_tot_cell.row
        num_daily_rows = total_row_idx - 2  # Daily rows are between Row 2 and Row (total_row_idx - 1)
        print(f"-> Phát hiện {num_daily_rows} dòng dữ liệu NTB cần cập nhật (từ dòng 2 đến {total_row_idx - 1}).")

        # Get the latest `num_daily_rows` dates from Data
        latest_dates = unique_data_dates[-num_daily_rows:] if len(unique_data_dates) >= num_daily_rows else unique_data_dates
        latest_dates_desc = sorted(latest_dates, key=lambda x: str(x).split(' ')[0] if x else "", reverse=True)

        print(f"-> Các ngày sẽ hiển thị trong báo cáo (mới đến cũ): {latest_dates_desc}")

        def to_vn_int_str(val):
            return f"{int(val):,}".replace(",", ".")

        def to_vn_pct_str(val):
            return f"{val * 100:.2f}".replace(".", ",") + "%"

        # Calculate daily rows (without Column A, i.e., columns B to N)
        daily_update_values = []
        parsed_ntb_data = [] # For in-memory report rendering

        for d in latest_dates_desc:
            df_d = df[df['Time'] == d]
            vol = df_d['Volume'].sum()
            hm = df_d['Hàng Mới Về Trong Ngày'].sum()
            gan = df_d['Sản Lượng Gán'].sum()
            chua_gan = df_d['Sản Lượng Chưa Gán'].sum()
            gtc = df_d['Sản Lượng Giao Thành Công'].sum()
            tra = df_d['Sản Lượng Chuyển Trả'].sum()
            ton = df_d['Sản Lượng Tồn'].sum()

            pct_gan = gan / vol if vol > 0 else 0.0
            pct_chua_gan = chua_gan / vol if vol > 0 else 0.0
            pct_gtc = gtc / vol if vol > 0 else 0.0
            pct_tra = tra / vol if vol > 0 else 0.0
            pct_ton = ton / vol if vol > 0 else 0.0

            # Row for updating columns B to N (13 elements)
            daily_update_values.append([
                d,
                to_vn_int_str(vol),
                to_vn_int_str(hm),
                to_vn_int_str(gan),
                to_vn_int_str(chua_gan),
                to_vn_int_str(gtc),
                to_vn_int_str(tra),
                to_vn_int_str(ton),
                to_vn_pct_str(pct_gan),
                to_vn_pct_str(pct_chua_gan),
                to_vn_pct_str(pct_gtc),
                to_vn_pct_str(pct_tra),
                to_vn_pct_str(pct_ton)
            ])

            # Full row for parsed_tq["ntb_data"] in memory (with Column A)
            parsed_ntb_data.append([
                'NTB',
                d,
                to_vn_int_str(vol),
                to_vn_int_str(hm),
                to_vn_int_str(gan),
                to_vn_int_str(chua_gan),
                to_vn_int_str(gtc),
                to_vn_int_str(tra),
                to_vn_int_str(ton),
                to_vn_pct_str(pct_gan),
                to_vn_pct_str(pct_chua_gan),
                to_vn_pct_str(pct_gtc),
                to_vn_pct_str(pct_tra),
                to_vn_pct_str(pct_ton)
            ])

        print("-> Pivot Table tự động cập nhật dữ liệu NTB từ dòng 2 đến dòng 9.")


        # Recalculate NTB total row
        def recalculate_ntb_total(ntb_data):
            total_vol = 0
            total_hm = 0
            total_gan = 0
            total_chua_gan = 0
            total_gtc = 0
            total_tra = 0
            total_ton = 0

            for r in ntb_data:
                if len(r) < 3 or not r[2]: continue
                try:
                    total_vol += int(round(float(str(r[2]).replace('.', '').replace(',', '.'))))
                except: pass
                try:
                    total_hm += int(round(float(str(r[3]).replace('.', '').replace(',', '.'))))
                except: pass
                try:
                    total_gan += int(round(float(str(r[4]).replace('.', '').replace(',', '.'))))
                except: pass
                try:
                    total_chua_gan += int(round(float(str(r[5]).replace('.', '').replace(',', '.'))))
                except: pass
                try:
                    total_gtc += int(round(float(str(r[6]).replace('.', '').replace(',', '.'))))
                except: pass
                try:
                    total_tra += int(round(float(str(r[7]).replace('.', '').replace(',', '.'))))
                except: pass
                try:
                    total_ton += int(round(float(str(r[8]).replace('.', '').replace(',', '.'))))
                except: pass

            pct_gan = total_gan / total_vol if total_vol > 0 else 0.0
            pct_chua_gan = total_chua_gan / total_vol if total_vol > 0 else 0.0
            pct_gtc = total_gtc / total_vol if total_vol > 0 else 0.0
            pct_tra = total_tra / total_vol if total_vol > 0 else 0.0
            pct_ton = total_ton / total_vol if total_vol > 0 else 0.0

            return [
                "Tổng cộng",
                "",
                to_vn_int_str(total_vol),
                to_vn_int_str(total_hm),
                to_vn_int_str(total_gan),
                to_vn_int_str(total_chua_gan),
                to_vn_int_str(total_gtc),
                to_vn_int_str(total_tra),
                to_vn_int_str(total_ton),
                to_vn_pct_str(pct_gan),
                to_vn_pct_str(pct_chua_gan),
                to_vn_pct_str(pct_gtc),
                to_vn_pct_str(pct_tra),
                to_vn_pct_str(pct_ton)
            ]

        new_total_row = recalculate_ntb_total(parsed_ntb_data)
        parsed_tq["ntb_data"] = parsed_ntb_data
        parsed_tq["ntb_total"] = new_total_row

        # Calculate only, do not statically update the sheet to prevent Pivot Table #REF! collisions
        print("-> NTB total row dynamically computed for the HTML report.")

        # Determine the latest target date
        if latest_dates_desc:
            latest_date_str = latest_dates_desc[0]
        else:
            latest_date_str = parsed_tq["ntb_data"][0][1]

        print(f"-> Ngày cập nhật mới nhất: {latest_date_str}")

        # Current time markers
        tz = pytz.timezone("Asia/Ho_Chi_Minh")
        now = datetime.now(tz)
        time_label = now.strftime('%H:%M')
        date_label = now.strftime('%d/%m/%Y')

        # Filter for the target date
        df_date = df[df['Time'] == latest_date_str].copy()
        print(f"-> Số dòng thô trên ngày {latest_date_str}: {len(df_date)}")

        # Aggregate AM stats dynamically to ensure alignment with the target date
        am_list = [row[0] for row in parsed_tq["am_data"]]
        print(f"📊 Đang tự động tính toán dữ liệu AM cho ngày {latest_date_str}...")
        calculated_am_data, calculated_am_total = aggregate_am_stats(df_date, am_list)
        parsed_tq["am_data"] = calculated_am_data
        parsed_tq["am_total"] = calculated_am_total

        # AM table is updated dynamically as a pivot table, no static update to sheet needed
        print("-> AM table dynamically computed for the HTML report.")


        # Group by Chi tiết and aggregate
        agg_df = df_date.groupby('Chi tiết').agg({
            'Time': 'first',
            'Volume': 'sum',
            'Hàng Mới Về Trong Ngày': 'sum',
            'Sản Lượng Gán': 'sum',
            'Sản Lượng Giao Thành Công': 'sum',
            'Sản Lượng Tồn': 'sum'
        }).reset_index()
        
        # Build AM mapping from Cơ cấu worksheet data in parsed_tq or read dynamically
        print("📖 Lấy danh sách AM từ tab Cơ cấu...")
        try:
            ws_cocau = None
            for sname in ["CoCauVung", "Cơ cấu", "cơ cấu"]:
                try:
                    ws_cocau = sh.worksheet(sname)
                    break
                except Exception:
                    pass
            cocau_rows = ws_cocau.get_all_values() if ws_cocau else []
            cocau_map = {}
            for row in cocau_rows[1:]:
                if len(row) >= 5:
                    bc_name = row[2].strip() # Column C - BC (Mapped name)
                    am_name = row[4].strip() # Column E - Am (AM name)
                    if bc_name and am_name:
                        cocau_map[bc_name] = am_name
        except Exception as e:
            print(f"⚠️ Không đọc được Cơ cấu cho AM mapping: {e}")
            cocau_map = {}

        # Calculate percentages and map AM name
        agg_df['%Gán'] = (agg_df['Sản Lượng Gán'] / agg_df['Volume']).fillna(0)
        agg_df['%GTC'] = (agg_df['Sản Lượng Giao Thành Công'] / agg_df['Volume']).fillna(0)
        agg_df['AM'] = agg_df['Chi tiết'].str.strip().map(cocau_map).fillna('')

        # Filter volume >= 300
        filtered_df = agg_df[agg_df['Volume'] >= 300].copy()
        print(f"-> Số bưu cục đạt sản lượng >= 300: {len(filtered_df)}")

        # Best 20
        best_20 = filtered_df.sort_values(by='%GTC', ascending=False).head(20)
        # Worst 20
        worst_20 = filtered_df.sort_values(by='%GTC', ascending=True).head(20)

    except Exception as e:
        print(f"❌ Lỗi khi xử lý dữ liệu bưu cục: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

    # 3. HTML Rendering using Local HTML & Playwright Element Screenshot
    print("-> Đang build HTML cho các bảng báo cáo...")
    html_overview = build_overview_html(parsed_tq, time_label, date_label)
    html_best = build_top_offices_html("TOP 20 BƯU CỤC %GTC TỐT NHẤT", best_20, is_best=True, date_str=latest_date_str)
    html_worst = build_top_offices_html("TOP 20 BƯU CỤC %GTC TỆ NHẤT", worst_20, is_best=False, date_str=latest_date_str)

    temp_tq_path = "temp_overview.html"
    temp_best_path = "temp_best.html"
    temp_worst_path = "temp_worst.html"

    with open(temp_tq_path, "w", encoding="utf-8") as f:
        f.write(html_overview)
    with open(temp_best_path, "w", encoding="utf-8") as f:
        f.write(html_best)
    with open(temp_worst_path, "w", encoding="utf-8") as f:
        f.write(html_worst)

    img_paths = {
        "overview": "table_overview.png",
        "best": "table_best_GTC.png",
        "worst": "table_worst_GTC.png"
    }

    print("-> Đang chụp ảnh các bảng bằng Playwright...")
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch()
            
            # Capture Overview
            page = browser.new_page()
            page.set_viewport_size({"width": 1460, "height": 1800})
            page.goto(f"file:///{os.path.abspath(temp_tq_path)}")
            page.wait_for_timeout(500)
            page.locator("#overview-container").screenshot(path=img_paths["overview"])
            
            # Capture Best
            page2 = browser.new_page()
            page2.set_viewport_size({"width": 1160, "height": 1300})
            page2.goto(f"file:///{os.path.abspath(temp_best_path)}")
            page2.wait_for_timeout(500)
            page2.locator("#capture-container").screenshot(path=img_paths["best"])
            
            # Capture Worst
            page3 = browser.new_page()
            page3.set_viewport_size({"width": 1160, "height": 1300})
            page3.goto(f"file:///{os.path.abspath(temp_worst_path)}")
            page3.wait_for_timeout(500)
            page3.locator("#capture-container").screenshot(path=img_paths["worst"])
            
            browser.close()
        print("✔️ Đã chụp thành công 3 ảnh bảng báo cáo.")
    except Exception as e:
        print(f"❌ Lỗi chụp ảnh Playwright: {e}")
        sys.exit(1)

    # 4. Broadcast to Telegram & GTalk
    print("📡 Đang gửi báo cáo qua Telegram và GTalk...")

    # Messages
    msg_overview_tele = f"📊 <b>BÁO CÁO VẬN HÀNH TỔNG QUAN</b>\n" \
                        f"⏱️ <b>Mốc cập nhật:</b> {time_label} ngày {date_label}\n" \
                        f"📈 <b>Vùng NTB ({latest_date_str}):</b>\n" \
                        f"  • Vol: <b>{format_vn_int(parsed_tq['ntb_data'][0][2])}</b> | GTC: <b>{format_vn_int(parsed_tq['ntb_data'][0][6])}</b>\n" \
                        f"  • %Gán: <b>{format_vn_pct(parsed_tq['ntb_data'][0][9])}</b> | %GTC: <b>{format_vn_pct(parsed_tq['ntb_data'][0][11])}</b> | %Tồn: <b>{format_vn_pct(parsed_tq['ntb_data'][0][13])}</b>"

    msg_overview_gtalk = msg_overview_tele

    msg_best_tele = f"🟢 <b>TOP 20 BƯU CỤC %GTC TỐT NHẤT</b>\n" \
                    f"⏱️ <b>Ngày:</b> {latest_date_str}\n" \
                    f"🏆 Bưu cục dẫn đầu: <b>{best_20.iloc[0]['Chi tiết'] if not best_20.empty else ''}</b> (%GTC: <b>{format_vn_pct(best_20.iloc[0]['%GTC']) if not best_20.empty else ''}</b>)"

    msg_best_gtalk = msg_best_tele

    msg_worst_tele = f"🔴 <b>TOP 20 BƯU CỤC %GTC TỆ NHẤT</b>\n" \
                     f"⏱️ <b>Ngày:</b> {latest_date_str}\n" \
                     f"⚠️ Bưu cục thấp nhất: <b>{worst_20.iloc[0]['Chi tiết'] if not worst_20.empty else ''}</b> (%GTC: <b>{format_vn_pct(worst_20.iloc[0]['%GTC']) if not worst_20.empty else ''}</b>)"

    msg_worst_gtalk = msg_worst_tele

    # Send Overview
    send_photo_telegram(img_paths["overview"], msg_overview_tele)
    send_photo_gtalk(img_paths["overview"], msg_overview_gtalk)
    time.sleep(2)

    # Send Best
    send_photo_telegram(img_paths["best"], msg_best_tele)
    send_photo_gtalk(img_paths["best"], msg_best_gtalk)
    time.sleep(2)

    # Send Worst
    send_photo_telegram(img_paths["worst"], msg_worst_tele)
    send_photo_gtalk(img_paths["worst"], msg_worst_gtalk)

    # Clean up local temporary files
    try:
        os.remove(temp_tq_path)
        os.remove(temp_best_path)
        os.remove(temp_worst_path)
        for p in img_paths.values():
            if os.path.exists(p):
                os.remove(p)
        print("🧹 Đã dọn dẹp các tệp tạm.")
    except Exception as e:
        print(f"⚠️ Cảnh báo: Không thể dọn dẹp một số tệp tạm: {e}")

    print("🎉 HOÀN THÀNH BÁO CÁO TỔNG QUAN THÀNH CÔNG!")

if __name__ == "__main__":
    main()
