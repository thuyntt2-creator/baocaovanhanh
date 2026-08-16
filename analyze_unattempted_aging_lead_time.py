# -*- coding: utf-8 -*-
"""
Script: analyze_unattempted_aging_lead_time.py
Analyzes the average sitting time (Tuổi đơn chưa giao TB) for active unattempted orders
across each AM.

Writes to a styled tab 'Độ trì trệ chưa giao' in Google Sheets,
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

def main():
    # Check active hours
    tz = pytz.timezone("Asia/Ho_Chi_Minh")
    current_hour = datetime.now(tz).hour
    bypass_time = len(sys.argv) > 1 and sys.argv[1] == "--force"
    if not bypass_time and not (7 <= current_hour <= 18):
        print(f"💤 Ngoài khung giờ hoạt động (7h - 18h). Hiện tại là {datetime.now(tz).strftime('%H:%M:%S')}. Script sẽ dừng.")
        sys.exit(0)

    print(f"🚀 Bắt đầu phân tích độ trì trệ chưa giao của AM lúc: {datetime.now(tz).strftime('%Y-%m-%d %H:%M:%S')}")

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
        ag_am_col = aging_header.index("am_name")
        ag_num_col = aging_header.index("num_deliver") if "num_deliver" in aging_header else -1
        ag_aging_col = aging_header.index("Aging") if "Aging" in aging_header else aging_header.index("aging")
    except ValueError as e:
        print(f"❌ Lỗi: Không tìm thấy cột bắt buộc trong tab Đơn giao aging trên 5 ngày. Chi tiết: {e}")
        sys.exit(1)

    # 2. Process active unattempted orders and track their aging
    success_keywords = ['đã giao/ chuyển trả thành công', 'đã giao/chuyển trả thành công', 'n/a', 'thành công']
    am_aging_data = {}

    for row in aging_data[1:]:
        if len(row) > max(ag_order_col, ag_bc_col, ag_id_bc_col, ag_am_col, ag_aging_col):
            order_code = row[ag_order_col].strip()
            if not order_code:
                continue

            status = lm_status.get(order_code, '#N/A')
            is_processed = status in ['#N/A', 'n/a'] or any(sk in status.lower() for sk in success_keywords)
            if is_processed:
                continue

            num_deliver_val = row[ag_num_col].strip() if (ag_num_col != -1 and len(row) > ag_num_col) else '0'
            try:
                num_attempts = int(float(num_deliver_val))
            except ValueError:
                num_attempts = 0

            # Only process unattempted orders
            if num_attempts == 0:
                raw_am = row[ag_am_col].strip()
                if not raw_am or raw_am == '#N/A' or raw_am == '':
                    bc_name = row[ag_bc_col].strip()
                    id_bc = row[ag_id_bc_col].strip()
                    am_name = cocau_map.get(id_bc, cocau_map.get(bc_name, "Không xác định"))
                else:
                    am_name = raw_am
                am_name = unicodedata.normalize('NFC', am_name)

                try:
                    aging_days = float(row[ag_aging_col].strip())
                except ValueError:
                    aging_days = 0.0

                if am_name not in am_aging_data:
                    am_aging_data[am_name] = []
                am_aging_data[am_name].append(aging_days)

    # Calculate average aging per AM
    am_sitting_leaderboard = []
    for am, aging_list in am_aging_data.items():
        if aging_list:
            avg_val = sum(aging_list) / len(aging_list)
            am_sitting_leaderboard.append({
                "am": am,
                "avg_days": avg_val,
                "count": len(aging_list)
            })

    # Sort descending by average aging days
    am_sitting_leaderboard = sorted(am_sitting_leaderboard, key=lambda x: x["avg_days"], reverse=True)

    # 3. Create or write to Độ trì trệ chưa giao tab
    print("📊 Cập nhật dữ liệu vào Google Sheets...")
    ws_name = "Độ trì trệ chưa giao"
    try:
        ws = sh.worksheet(ws_name)
        ws.clear()
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=ws_name, rows="200", cols="5")

    now_str = datetime.now(tz).strftime('%Y-%m-%d %H:%M:%S')
    grid_values = []
    grid_values.append(["BẢNG THEO DÕI ĐỘ TRÌ TRỆ CHƯA GIAO CỦA AM", "", ""])
    grid_values.append([f"Cập nhật lúc: {now_str}", "", ""])
    grid_values.append([""]) # Spacer row
    grid_values.append(["AM", "Tuổi đơn chưa giao TB (ngày)", "Số đơn chưa giao"])

    for item in am_sitting_leaderboard:
        grid_values.append([
            item["am"],
            f"{item['avg_days']:.1f}",
            str(item["count"])
        ])

    sheet_requests = []
    sheet_requests.append(merge_request(ws.id, 0, 1, 0, 3))
    sheet_requests.append(merge_request(ws.id, 1, 2, 0, 3))

    sheet_requests.append(cell_format_request(ws.id, 0, 1, 0, 3, {
        "textFormat": {"bold": True, "fontSize": 14, "fontFamily": "Arial", "foregroundColor": make_color("#7F1D1D")},
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "MIDDLE"
    }))
    sheet_requests.append(cell_format_request(ws.id, 1, 2, 0, 3, {
        "textFormat": {"fontSize": 10, "fontFamily": "Arial", "italic": True, "foregroundColor": make_color("#555555")},
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "MIDDLE"
    }))
    sheet_requests.append(cell_format_request(ws.id, 3, 4, 0, 3, {
        "backgroundColor": make_color("#7F1D1D"),
        "textFormat": {"bold": True, "fontSize": 10, "fontFamily": "Arial", "foregroundColor": make_color("#FFFFFF")},
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "MIDDLE"
    }))

    sheet_requests.append(row_height_request(ws.id, 0, 1, 35))
    sheet_requests.append(row_height_request(ws.id, 1, 2, 24))
    sheet_requests.append(row_height_request(ws.id, 3, 4, 30))

    for am_idx, item in enumerate(am_sitting_leaderboard):
        row_idx = 4 + am_idx
        avg_days = item["avg_days"]
        if avg_days > 15.0:
            bg = "#FFCDD2"
            fg = "#B71C1C"
        elif avg_days > 10.0:
            bg = "#FFF9C4"
            fg = "#78350F"
        else:
            bg = "#C8E6C9"
            fg = "#1B5E20"
            
        sheet_requests.append(cell_format_request(ws.id, row_idx, row_idx+1, 1, 2, {
            "backgroundColor": make_color(bg),
            "textFormat": {"foregroundColor": make_color(fg), "bold": True, "fontFamily": "Arial", "fontSize": 9},
            "horizontalAlignment": "CENTER",
            "verticalAlignment": "MIDDLE"
        }))
        sheet_requests.append(cell_format_request(ws.id, row_idx, row_idx+1, 0, 1, {
            "backgroundColor": make_color("#FADBD8"),
            "textFormat": {"bold": True, "fontFamily": "Arial", "fontSize": 9},
            "horizontalAlignment": "LEFT",
            "verticalAlignment": "MIDDLE"
        }))
        sheet_requests.append(cell_format_request(ws.id, row_idx, row_idx+1, 2, 3, {
            "backgroundColor": make_color("#F2F4F4"),
            "textFormat": {"fontFamily": "Arial", "fontSize": 9},
            "horizontalAlignment": "CENTER",
            "verticalAlignment": "MIDDLE"
        }))
        sheet_requests.append(row_height_request(ws.id, row_idx, row_idx+1, 26))

    sheet_requests.append(border_request(ws.id, 3, 4 + len(am_sitting_leaderboard), 0, 3))
    sheet_requests.append(col_width_request(ws.id, 0, 1, 180))
    sheet_requests.append(col_width_request(ws.id, 1, 2, 200))
    sheet_requests.append(col_width_request(ws.id, 2, 3, 140))

    # Update sheet content
    ws.update(range_name=f"A1:C{len(grid_values)}", values=grid_values, value_input_option="USER_ENTERED")
    sh.batch_update({"requests": sheet_requests})
    print("✔️ Đã cập nhật xong sheet 'Độ trì trệ chưa giao'.")

    # 4. Generate visual HTML table & capture screenshot
    print("📸 Renderer: Vẽ bảng HTML chưa giao và chụp hình...")
    
    tbody_sitting_rows = ""
    for item in am_sitting_leaderboard:
        avg_val = item["avg_days"]
        if avg_val > 15.0:
            badge_class = "rate-red"
        elif avg_val > 10.0:
            badge_class = "rate-yellow"
        else:
            badge_class = "rate-green"
            
        tbody_sitting_rows += f"""
        <tr>
            <td class="am-name">{item['am']}</td>
            <td>
                <div class="rate-badge {badge_class}">{avg_val:.1f} ngày</div>
            </td>
            <td style="font-weight: 600; color: #475569;">{item['count']} đơn</td>
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
            width: 700px;
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
            font-size: 22px;
            color: #7f1d1d;
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
            padding: 12px 6px;
            border-bottom: 1px solid #f1f5f9;
            vertical-align: middle;
            font-size: 14px;
        }}
        .am-name {{
            font-weight: 600;
            color: #0f172a;
            text-align: left;
            padding-left: 12px;
            background-color: #f8fafc;
        }}
        .rate-badge {{
            display: inline-block;
            padding: 5px 12px;
            border-radius: 6px;
            font-size: 13px;
            font-weight: 700;
            min-width: 80px;
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
    </style>
    </head>
    <body>
    <div id="table-container">
        <div class="header">
            <h2>Bảng độ trì trệ chưa giao của AM</h2>
            <div class="subtitle">Tuổi đơn (Aging) trung bình của các đơn chưa có lần đi giao nào (Cập nhật: {now_str})</div>
        </div>
        <table>
            <thead>
                <tr>
                    <th style="text-align: left; padding-left: 12px; width: 280px;">AM</th>
                    <th style="width: 220px;">Tuổi đơn chưa giao TB</th>
                    <th style="width: 200px;">Số đơn chưa giao</th>
                </tr>
            </thead>
            <tbody>
                {tbody_sitting_rows}
            </tbody>
        </table>
    </div>
    </body>
    </html>
    """

    temp_html_path = os.path.join(BASE_DIR, "temp_am_sitting_analysis.html")
    with open(temp_html_path, "w", encoding="utf-8") as f:
        f.write(html_content)

    output_image_path = os.path.join(BASE_DIR, "table_am_sitting_analysis.png")
    
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

    # Calculate Top 5 text for caption
    sitting_lines = []
    sitting_lines.append("⏳ <b>Độ trì trệ chưa giao trung bình của AM (Tuổi đơn TB):</b>")
    for idx, item in enumerate(am_sitting_leaderboard[:5]):
        sitting_lines.append(f"• AM <b>{item['am']}</b>: tồn TB <b>{item['avg_days']:.1f}</b> ngày ({item['count']} đơn)")
    sitting_text = "\n".join(sitting_lines)

    # 5. Gửi sang GTalk
    print("📡 Đang gửi ảnh báo cáo phân tích độ trì trệ sang GTalk group...")
    
    caption_text = (
        f"🚨 <b>BÁO CÁO PHÂN TÍCH ĐỘ TRÌ TRỆ CHƯA GIAO CỦA AM</b>\n"
        f"📅 Ngày: {datetime.now(tz).strftime('%d-%m-%Y')} | ⏱️ Mốc cập nhật: {datetime.now(tz).strftime('%H:%M')}\n"
        f"========================\n"
        f"{sitting_text}\n"
        f"========================\n"
        f"💡 <b>Lưu ý:</b> Độ trì trệ (Tuổi đơn TB) càng cao có nghĩa AM đó đang bỏ quên đơn chưa giao lâu nhất. Cần xử lý dứt điểm đơn hàng để giảm độ trì trệ.\n"
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
        "Metadata": json.dumps({"width": 700, "height": 600}),
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
                                    "items": [{"image": {"fileId": file_id, "width": 700, "height": 600}}]
                                }
                            },
                            "oaToken": GTALK_OA_TOKEN
                        }
                        resp_send = requests.post("https://mbff.ghn.vn/api/gtalk/send-message", json=send_payload)
                        if resp_send.status_code == 200:
                            print("✅ Đã gửi báo cáo độ trì trệ lên GTalk group chat thành công!")
                        else:
                            print(f"❌ Lỗi gửi tin nhắn GTalk: {resp_send.status_code} - {resp_send.text}")
    
    # Gửi sang Telegram
    tele_token = "8570130113:AAGXRiUaKBknVpgtm1_i9ZA47JRjAXmB21M"
    tele_chat = "-5058464865"
    print("📡 Đang gửi ảnh báo cáo phân tích độ trì trệ sang Telegram group...")
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
