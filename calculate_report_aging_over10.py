import sys, io
import os
import requests
import asyncio
import json
from telethon import TelegramClient
from datetime import datetime, timezone, timedelta
import re
import pandas as pd
import gspread
import unicodedata
from google.oauth2.service_account import Credentials
from playwright.sync_api import sync_playwright

# Fix encoding cho Task Scheduler/Command Prompt
os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except AttributeError:
    pass

# Cấu hình và hằng số
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
SNAPSHOT_FILE = os.path.join(BASE_DIR, 'snapshot_aging_over10.json')

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

# Cấu hình của Thủy (Telegram Bot)
API_ID = 33980755
API_HASH = '27cb91d2027884b61393c554a4439dff'
PHONE = '+84368644943'
BOT_NAME = 'ghn_staff_bot'

# Cấu hình Google Sheet
SHEET_KEY = '1vCxSTNgSpO9ETvVRElGyuGc7lnx7LxLRhAB4-lJMHLU'
SHEET_NAME = 'data LM'
TEMP_EXCEL = os.path.join(BASE_DIR, 'temp_ghn.xlsx')
SESSION_FILE = os.path.join(BASE_DIR, 'session_thuy')

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def get_gspread_client(sheet_key=SHEET_KEY):
    if os.path.exists(JSON_FILE):
        try:
            creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
            gc = gspread.authorize(creds)
            if sheet_key:
                gc.open_by_key(sheet_key)
            return gc
        except Exception as e:
            print(f"⚠️ Service account ({JSON_FILE}) không có quyền truy cập: {e}. Đang chuyển sang authorized_user.json...")
            
    auth_user_file = os.path.join(BASE_DIR, 'authorized_user.json')
    if os.path.exists(auth_user_file):
        from google.oauth2.credentials import Credentials as UserCredentials
        creds = UserCredentials.from_authorized_user_file(auth_user_file, scopes=scopes)
        gc = gspread.authorize(creds)
        if sheet_key:
            gc.open_by_key(sheet_key)
        return gc
    raise PermissionError("Không thể xác thực Google Sheets bằng credentials.json hoặc authorized_user.json")


def get_ghn_token():
    config_path = os.path.join(BASE_DIR, 'ghn_config.json')
    default_token = "Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJvcmdDb2RlIjoiZ2huZXhwcmVzcyIsInBhcnRuZXJDb2RlIjoiIiwic2VlZCI6NTc5MDk4NTMzOTUyNjcxMjEwOSwic3NvSWQiOiIzMDY2MDIxIiwidXNlcklkIjoiNjRlMmUyYzI2MmNhZDU1YzZiODRlZTBkIn0.X52JxA4BJLZRY6dAgo793Of-RTNtORnLw0dpZnTlXtY"
    if os.path.exists(config_path):
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                return data.get("ghn_token", default_token)
        except Exception:
            pass
    return default_token

async def download_report_data():
    """Kích hoạt bot GHN và tải file Excel từ Telegram, cập nhật tab 'data LM'"""
    print(f"🚀 Bắt đầu tải dữ liệu từ GHN lúc: {datetime.now().strftime('%H:%M:%S')}")

    # A. Kích hoạt GHN xuất file
    headers = {
        'authorization': get_ghn_token(),
        'content-type': 'application/json',
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
    }

    id_list = ["1896","2309","2357","2399","2502","20143000","20144000","20150000","20269000","20316000","20320000","20367000","20495000","20499000","20543000","20558000","20588000","20590000","20591000","20633000","20648000","20663000","20669000","20670000","20673000","20687000","20691000","20694000","20784000","20785000","20849000","20942000","21046000","21065000","21067000","21126000","21150000","21290000","21294000","21298000","21320000","21331000","21364000","21377000","21396000","21403000","21456000","21458000","21468000","21477000","21479000","21516000","21537000","21569000","21594000","21659000","21687000","21688000","21722000","22048000","22051000","22052000","22086000","22116000","22119000","22153000","22154000","22155000","22222000","22242000","22255000","22263000","22299000","22312000","22329000","22356000","22357000","20495001","22357001","22363000","22375000","20591001","22389000","22394000","21366001","22425000","22452000","21377001","22483000","20269003","20785001","20590004","22549000","22483001","20663002","21126003","22255001","20663003","20942002","22704000","22051001","20785002","22746000","22759000","22774000","20785003","22793000","22048001","20673001","22830000","20673002","22861000","22861001","20591002","22915000","22934000"]
    json_data = {
        "hub_ids": id_list, 
        "is_all_hub": True,
        "type": 1,
        "customer_id": None,
        "ward_code": None,
        "is_count": False
    }

    print("📡 Gửi yêu cầu xuất file cho TẤT CẢ bưu cục...")
    trigger_time = datetime.now(timezone.utc) - timedelta(minutes=60)
    
    response = requests.post(
        'https://nhanh-api.ghn.vn/api/core/oss/v1/report/export-backlog-detail', 
        headers=headers, 
        json=json_data
    )

    if response.status_code == 200:
        print("✅ GHN đã nhận lệnh tổng! Đang đợi Telegram Bot báo link...")
    elif response.status_code == 400 and ("Yêu cầu đang được xử lý" in response.text or "1000" in response.text):
        print("⚠️ GHN báo yêu cầu đang được xử lý (vừa gửi yêu cầu trước đó). Tiếp tục lắng nghe Telegram Bot để nhận link...")
    else:
        print(f"❌ Lỗi kích hoạt GHN API: {response.status_code} - {response.text}")
        sys.exit(1)

    # B. Kết nối Telegram
    client = TelegramClient(SESSION_FILE, API_ID, API_HASH)
    await client.connect()
    if not await client.is_user_authorized():
        print("❌ Lỗi: Session Telegram chưa được xác thực.")
        await client.disconnect()
        sys.exit(1)

    print("⏳ Đang lắng nghe tin nhắn từ Bot GHN...")
    file_url = None
    for _ in range(60): 
        await asyncio.sleep(10)
        async for message in client.iter_messages(BOT_NAME, limit=5):
            if message.date < trigger_time: 
                continue
            if message.text and 'online-gateway.ghn.vn' in message.text and 'Bao_cao_ton_dong_lay_giao_tra' in message.text:
                urls = re.findall(r'https?://online-gateway\.ghn\.vn[^\s)]+', message.text)
                if urls:
                    file_url = urls[0]
                    break
        if file_url: 
            break

    await client.disconnect()

    # C. Tải file về và đẩy lên Google Sheets
    if file_url:
        print(f"✅ Đã tìm thấy link tải: {file_url}")
        headers_dl = {'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
        res = None
        for attempt in range(5):
            try:
                res = requests.get(file_url, headers=headers_dl, timeout=30)
                if res.status_code == 200:
                    break
            except Exception as e:
                print(f"⚠️ Lần {attempt+1}/5 tải file từ GHN gặp sự cố ({e}). Đang thử lại sau 3s...")
                await asyncio.sleep(3)

        if not res or res.status_code != 200:
            print("❌ Lỗi: Không thể tải file Excel từ GHN Gateway sau 5 lần thử.")
            sys.exit(1)

        with open(TEMP_EXCEL, 'wb') as f:
            f.write(res.content)

        print("📊 Đọc dữ liệu Excel và cập nhật lên Google Sheets...")
        df = pd.read_excel(TEMP_EXCEL).fillna("")
        total_rows = len(df)
        print(f"📦 Đã tải về thành công: {total_rows} đơn hàng!")

        gc_client = get_gspread_client(SHEET_KEY)
        sh = gc_client.open_by_key(SHEET_KEY)
        worksheet = sh.worksheet(SHEET_NAME)

        worksheet.clear()
        worksheet.update([df.columns.values.tolist()] + df.values.tolist())
        print("✔️ Đã cập nhật xong dữ liệu thô vào tab 'data LM'.")
    else:
        print("❌ Lỗi: Hết thời gian chờ mà Bot GHN không phản hồi link.")
        sys.exit(1)

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

def get_group_over10(aging_val, group_val):
    try:
        age = float(aging_val)
        if age > 15.0:
            return 'Trên 15 ngày'
        elif age > 10.0:
            return '10 - 15 ngày'
    except ValueError:
        pass
    val = str(group_val).strip().lower()
    if any(k in val for k in ['(f)', '(g)', '(h)', '(i)', '(j)']):
        return '10 - 15 ngày'
    elif '(k)' in val:
        return 'Trên 15 ngày'
    return None

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

def run_calculations():
    """Đọc dữ liệu, tính toán gán đơn và PIVOT, cập nhật sheet trực tiếp và gửi thông báo"""
    print(f"🔄 Bắt đầu chạy quy trình tính toán lúc: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    
    # Kết nối Google Sheets
    gc_client = get_gspread_client(SHEET_KEY)
    sh = gc_client.open_by_key(SHEET_KEY)
    
    # 1. Đọc dữ liệu từ sheet "data LM"
    print("📖 Đọc dữ liệu từ sheet 'data LM'...")
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
        print(f"❌ Lỗi: Không tìm thấy cột cần thiết trong 'data LM'. Chi tiết: {e}")
        sys.exit(1)
        
    # Map Mã đơn hàng -> Trạng thái
    lm_status = {}
    for row in lm_data[2:]:
        if len(row) > max(order_col_idx, status_col_idx):
            m_don = row[order_col_idx].strip()
            t_thai = row[status_col_idx].strip()
            if m_don:
                lm_status[m_don] = t_thai
                
    # 2. Đọc bảng Cơ cấu
    print("📖 Đọc bảng Cơ cấu...")
    ws_cocau = None
    for sname in ["CoCauVung", "Cơ cấu", "cơ cấu"]:
        try:
            ws_cocau = sh.worksheet(sname)
            break
        except Exception:
            pass
    if not ws_cocau:
        print("❌ Không tìm thấy tab CoCauVung hoặc Cơ cấu trong spreadsheet.")
        sys.exit(1)
        
    cocau_data = ws_cocau.get_all_values()
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

    # 3. Đọc dữ liệu từ sheet gốc "Đơn giao aging trên 5 ngày"
    print("📖 Đọc dữ liệu từ sheet 'Đơn giao aging trên 5 ngày'...")
    ws_aging = sh.worksheet("Đơn giao aging trên 5 ngày")
    aging_data = ws_aging.get_all_values()
    
    if len(aging_data) < 1 or (len(aging_data) > 0 and (aging_data[0][0].startswith('#') or ("order_code" not in aging_data[0] and "mã đơn" not in aging_data[0]))):
        print("⚠️ Sheet 'Đơn giao aging trên 5 ngày' bị lỗi IMPORTRANGE (#REF!/#ERROR!). Tự động lấy dữ liệu trực tiếp từ Sheet nguồn...")
        try:
            sh_master = gc_client.open_by_key('1WCzgao34cA_SttyB9ytHfE1qKTNl_3iFqDbEfw3lbyU')
            ws_master_aging = sh_master.worksheet("Đơn giao aging trên 5 ngày")
            aging_data = ws_master_aging.get_all_values()
        except Exception as e:
            print(f"❌ Lỗi khi đọc Sheet nguồn: {e}")

    if len(aging_data) < 1:
        print("❌ Lỗi: Sheet 'Đơn giao aging trên 5 ngày' trống.")
        sys.exit(1)
        
    aging_header = aging_data[0]
    
    ag_order_idx = aging_header.index("order_code") if "order_code" in aging_header else aging_header.index("mã đơn")
    ag_bc_idx = aging_header.index("bc")
    ag_id_bc_idx = aging_header.index("id_bc")
    ag_aging_idx = aging_header.index("Aging") if "Aging" in aging_header else aging_header.index("aging")
    ag_group_idx = aging_header.index("Nhóm BL")
    ag_am_idx = aging_header.index("am_name")
    ag_tinh_idx = aging_header.index("tinh") if "tinh" in aging_header else aging_header.index("tỉnh")

    # 4. CẬP NHẬT TRẠNG THÁI GÁN VÀ LƯỢT GÁN (Logic từ update_aging_assignments.py)
    print("✍️ Đối chiếu và tính toán lượt gán đơn...")
    existing_tracking = {}
    try:
        ws_push = sh.worksheet("PUSH REGION")
        push_data = ws_push.get_all_values()
        if len(push_data) > 1:
            push_header = push_data[0]
            try:
                p_order_idx = push_header.index("Mã đơn hàng")
                p_status_idx = push_header.index("Trạng thái hiện tại")
                p_count_idx = push_header.index("Lượt gán")
                p_prev_idx = push_header.index("Trạng thái trước đó")
            except ValueError:
                p_order_idx, p_status_idx, p_count_idx, p_prev_idx = 0, 4, 6, 7
            
            for row in push_data[1:]:
                if len(row) > max(p_order_idx, p_status_idx, p_count_idx, p_prev_idx):
                    m_don = row[p_order_idx].strip()
                    if m_don:
                        existing_tracking[m_don] = {
                            "Trạng thái hiện tại": row[p_status_idx].strip(),
                            "Lượt gán": int(row[p_count_idx].strip()) if row[p_count_idx].strip().isdigit() else 0,
                            "Trạng thái trước đó": row[p_prev_idx].strip()
                        }
    except gspread.exceptions.WorksheetNotFound:
        print("   Tạo mới tab 'PUSH REGION'...")
        ws_push = sh.add_worksheet(title="PUSH REGION", rows="1000", cols="15")
        
    updated_rows = []
    push_headers = [
        'Mã đơn hàng', 'Tên bưu cục hiện tại', 'Tên tỉnh', 'AM', 
        'Trạng thái hiện tại', 'Có chuyến đi giao', 'Lượt gán', 
        'Trạng thái trước đó', 'Cập nhật lúc'
    ]
    
    now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    new_assignments_count = 0
    
    for row in aging_data[1:]:
        if len(row) > max(ag_order_idx, ag_bc_idx, ag_tinh_idx, ag_am_idx):
            m_don = row[ag_order_idx].strip()
            bc = row[ag_bc_idx].strip()
            tinh = row[ag_tinh_idx].strip()
            id_bc = row[ag_id_bc_idx].strip()
            raw_am = row[ag_am_idx].strip()
            if not raw_am or raw_am == '#N/A' or raw_am == '':
                am = cocau_map.get(id_bc, cocau_map.get(bc, "Không xác định"))
            else:
                am = raw_am
            am = unicodedata.normalize('NFC', am)
                
            if not m_don:
                continue
                
            status_now = lm_status.get(m_don, "đã giao/ chuyển trả thành công")
            
            if m_don in existing_tracking:
                prev_status = existing_tracking[m_don]["Trạng thái trước đó"]
                luot_gan = existing_tracking[m_don]["Lượt gán"]
            else:
                prev_status = ""
                luot_gan = 0
                
            is_assigned_now = status_now in ["Đang có chuyến đi trong ngày", "Đã có chuyến đi trong ngày"]
            was_assigned_prev = prev_status in ["Đang có chuyến đi trong ngày", "Đã có chuyến đi trong ngày"]
            
            if is_assigned_now and not was_assigned_prev:
                luot_gan += 1
                new_assignments_count += 1
                
            has_trip = "Có" if is_assigned_now else "Không"
            
            updated_rows.append([
                m_don, bc, tinh, am, 
                status_now, has_trip, str(luot_gan), 
                status_now,
                now_str
            ])
            
    print(f"   Hoàn thành đối chiếu. Phát hiện {new_assignments_count} lượt gán mới.")
    ws_push.clear()
    ws_push.update([push_headers] + updated_rows)
    print("✅ Đã cập nhật xong tab 'PUSH REGION'.")
    
    # Cập nhật thống kê vào sheet "Lượt gán"
    counts = {0: 0, 1: 0, 2: 0, 3: 0, "4+": 0}
    for row in updated_rows:
        cnt = int(row[6])
        if cnt in [0, 1, 2, 3]:
            counts[cnt] += 1
        else:
            counts["4+"] += 1
            
    summary_rows = [
        ["0", counts[0]],
        ["1", counts[1]],
        ["2", counts[2]],
        ["3", counts[3]],
        ["4+", counts["4+"]],
        ["Tổng cộng", len(updated_rows)]
    ]
    
    try:
        ws_summary = sh.worksheet("Lượt gán")
    except gspread.exceptions.WorksheetNotFound:
        ws_summary = sh.add_worksheet(title="Lượt gán", rows="10", cols="2")
        
    ws_summary.clear()
    ws_summary.update([["Lượt gán", "Số lượng đơn"]] + summary_rows)
    print("✅ Đã cập nhật xong tab 'Lượt gán'.")

    # 5. Phân tích các đơn hàng TRÊN 10 NGÀY
    all_over10_orders = []
    active_over10_orders = []
    
    for row in aging_data[1:]:
        if len(row) > max(ag_order_idx, ag_bc_idx, ag_id_bc_idx, ag_group_idx, ag_am_idx, ag_aging_idx):
            order_code = row[ag_order_idx].strip()
            if not order_code:
                continue
                
            group_val = row[ag_group_idx].strip()
            aging_val = row[ag_aging_idx].strip()
            
            group = get_group_over10(aging_val, group_val)
            if not group:
                # Không thuộc nhóm >10 ngày, bỏ qua
                continue
                
            status = lm_status.get(order_code, '#N/A')
            is_processed = status in ['#N/A', 'đã giao/ chuyển trả thành công', 'đã giao/chuyển trả thành công', 'n/a', 'thành công']
            
            bc_name = row[ag_bc_idx].strip()
            id_bc = row[ag_id_bc_idx].strip()
            
            raw_am_name = row[ag_am_idx].strip()
            if not raw_am_name or raw_am_name == '#N/A' or raw_am_name == '':
                am_name = cocau_map.get(id_bc, cocau_map.get(bc_name, "Không xác định"))
            else:
                am_name = raw_am_name
            am_name = unicodedata.normalize('NFC', am_name)
                
            order_info = {
                'order_code': order_code,
                'bc': bc_name,
                'id_bc': id_bc,
                'am': am_name,
                'group': group,
                'status': status,
                'is_processed': is_processed,
                'row_raw': row
            }
            
            all_over10_orders.append(order_info)
            if not is_processed:
                active_over10_orders.append(order_info)

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
            current_time = datetime.now().strftime("%H:%M")
            today_str = datetime.now().strftime("%Y-%m-%d")
            today_key = datetime.now().strftime("%Y%m%d")
    else:
        current_time = datetime.now().strftime("%H:%M")
        today_str = datetime.now().strftime("%Y-%m-%d")
        today_key = datetime.now().strftime("%Y%m%d")

    am_names = sorted(list(set([o['am'] for o in all_over10_orders])))
    group_labels = ['10 - 15 ngày', 'Trên 15 ngày']

    # Thống kê chi tiết per AM
    # pivot_map[am] = { '10 - 15 ngày': active, 'Trên 15 ngày': active, 'active_count': active_total, 'assigned': active_assigned, 'unassigned': active_unassigned, 'processed': processed_total, 'total_orders': total }
    pivot_map = {}
    for am in am_names:
        pivot_map[am] = {
            '10 - 15 ngày': 0, 
            'Trên 15 ngày': 0, 
            'active_count': 0, 
            'assigned': 0, 
            'unassigned': 0, 
            'processed': 0, 
            'total_orders': 0
        }
        
    for o in all_over10_orders:
        am = o['am']
        group = o['group']
        status = o['status']
        
        pivot_map[am]['total_orders'] += 1
        if o['is_processed']:
            pivot_map[am]['processed'] += 1
        else:
            pivot_map[am]['active_count'] += 1
            pivot_map[am][group] += 1
            is_assigned = status in ["Đang có chuyến đi trong ngày", "Đã có chuyến đi trong ngày"]
            if is_assigned:
                pivot_map[am]['assigned'] += 1
            else:
                pivot_map[am]['unassigned'] += 1

    # Thống kê chi tiết per BC
    bc_stats = {}
    for o in all_over10_orders:
        bc = o['bc']
        am = o['am']
        group = o['group']
        status = o['status']
        
        if bc not in bc_stats:
            bc_stats[bc] = {
                'am': am, 
                '10 - 15 ngày': 0, 
                'Trên 15 ngày': 0, 
                'active_count': 0, 
                'assigned': 0, 
                'unassigned': 0, 
                'processed': 0, 
                'total_orders': 0
            }
            
        bc_stats[bc]['total_orders'] += 1
        if o['is_processed']:
            bc_stats[bc]['processed'] += 1
        else:
            bc_stats[bc]['active_count'] += 1
            bc_stats[bc][group] += 1
            is_assigned = status in ["Đang có chuyến đi trong ngày", "Đã có chuyến đi trong ngày"]
            if is_assigned:
                bc_stats[bc]['assigned'] += 1
            else:
                bc_stats[bc]['unassigned'] += 1

    sorted_bcs = sorted(bc_stats.items(), key=lambda x: x[1]['active_count'], reverse=True)
    top5_bcs = sorted_bcs[:5]
    top5_bc_names = [x[0] for x in top5_bcs]

    current_am_active_totals = {am: pivot_map[am]['active_count'] for am in am_names}
    current_bc_active_totals = {bc: stats['active_count'] for bc, stats in bc_stats.items()}

    # 6. Quản lý Snapshot Lịch sử
    print("📝 Cập nhật lịch sử mốc thời gian cho Đơn >10 Ngày...")
    state = {"last_updated_date": "", "history": [], "daily_snapshots": {}}
    if os.path.exists(SNAPSHOT_FILE):
        try:
            with open(SNAPSHOT_FILE, "r", encoding="utf-8") as f:
                state = json.load(f)
        except:
            pass

    # Phục hồi lịch sử từ PIVOT nếu cần thiết
    try:
        print("🔍 Đang đọc lịch sử từ sheet PIVOT hiện tại...")
        ws_pivot = sh.worksheet("PIVOT")
        rows = ws_pivot.get_all_values()
        header_row_idx = -1
        for idx, row in enumerate(rows):
            if len(row) > 0 and "Đơn aging >10 ngày" in row[0] and "hằng ngày" in row[0]:
                header_row_idx = idx + 1
                break
        if header_row_idx != -1 and header_row_idx < len(rows):
            header_row = rows[header_row_idx]
            col_dates = {}
            current_year = datetime.now().year
            for col_idx in range(1, len(header_row)):
                cell_text = header_row[col_idx].strip()
                if not cell_text:
                    continue
                match = re.search(r'\((\d{2})/(\d{2})\)', cell_text)
                if match:
                    d_day, d_month = match.groups()
                    date_key = f"{current_year}{d_month}{d_day}"
                    col_dates[col_idx] = date_key
            
            for row_idx in range(header_row_idx + 1, len(rows)):
                row = rows[row_idx]
                if not row or len(row) == 0:
                    continue
                am_name = row[0].strip()
                if not am_name or am_name == "TỔNG" or am_name.upper() == "TOTAL":
                    break
                for col_idx, date_key in col_dates.items():
                    if date_key != today_key and col_idx < len(row):
                        val_str = row[col_idx].strip()
                        if val_str:
                            first_line = val_str.split('\n')[0].strip()
                            num_match = re.match(r'^\d+', first_line)
                            if num_match:
                                num_val = int(num_match.group())
                                if "daily_snapshots" not in state:
                                    state["daily_snapshots"] = {}
                                if date_key not in state["daily_snapshots"]:
                                    state["daily_snapshots"][date_key] = {"totals": {}, "grandTotal": 0}
                                if am_name not in state["daily_snapshots"][date_key]["totals"]:
                                    state["daily_snapshots"][date_key]["totals"][am_name] = num_val
            
            for dk in state.get("daily_snapshots", {}):
                totals = state["daily_snapshots"][dk]["totals"]
                state["daily_snapshots"][dk]["grandTotal"] = sum(totals.values())
            print("✅ Đã khôi phục dữ liệu lịch sử từ tab PIVOT thành công.")
    except Exception as e:
        print(f"⚠️ Cảnh báo: Không thể khôi phục lịch sử từ PIVOT: {e}")

    # Check new day
    if state.get("last_updated_date") != today_str:
        # Lưu mốc sáng ngày hôm trước
        prev_date = state.get("last_updated_date")
        if prev_date and len(state["history"]) > 0:
            prev_key = prev_date.replace("-", "")
            morning_snap = state["history"][0]
            state["daily_snapshots"][prev_key] = {
                "totals": morning_snap["totals"],
                "grandTotal": sum(morning_snap["totals"].values())
            }
        state["history"] = []
        state["last_updated_date"] = today_str

    current_snap = {
        "time": current_time,
        "totals": current_am_active_totals,
        "bcTotals": current_bc_active_totals
    }

    if len(state["history"]) == 0:
        # Nếu chạy lần đầu tiên trong ngày vào buổi chiều (e.g. sau 9h00 sáng), 
        # tự động tạo mốc 7h30 sáng giả lập bằng cách sử dụng tổng đơn của ngày (active + processed).
        hour_now = datetime.now().hour
        if hour_now >= 9 or custom_date_str:
            current_am_total_orders = {am: pivot_map[am]['total_orders'] for am in am_names}
            current_bc_total_orders = {bc: stats['total_orders'] for bc, stats in bc_stats.items()}
            morning_snap = {
                "time": "07:30",
                "totals": current_am_total_orders,
                "bcTotals": current_bc_total_orders
            }
            state["history"].append(morning_snap)
            state["history"].append(current_snap)
            state["daily_snapshots"][today_key] = {
                "totals": current_am_total_orders,
                "grandTotal": sum(current_am_total_orders.values())
            }
        else:
            state["history"].append(current_snap)
            state["daily_snapshots"][today_key] = {
                "totals": current_am_active_totals,
                "grandTotal": sum(current_am_active_totals.values())
            }
    elif len(state["history"]) == 1:
        state["history"].append(current_snap)
    else:
        state["history"][1] = current_snap

    with open(SNAPSHOT_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)

    history = state["history"]
    history_len = len(history)

    # 7. Ghi dữ liệu và format sheet PIVOT
    print("📊 Cập nhật dữ liệu PIVOT...")
    ws_pivot = sh.worksheet("PIVOT")
    ws_pivot.clear()

    # Xóa format cũ
    clear_format_req = {
        "repeatCell": {
            "range": {"sheetId": ws_pivot.id, "startRowIndex": 0, "endRowIndex": 150, "startColumnIndex": 0, "endColumnIndex": 25},
            "cell": {"userEnteredFormat": {}},
            "fields": "userEnteredFormat"
        }
    }
    unmerge_req = {
        "unmergeCells": {
            "range": {"sheetId": ws_pivot.id, "startRowIndex": 0, "endRowIndex": 150, "startColumnIndex": 0, "endColumnIndex": 25}
        }
    }
    requests = [clear_format_req, unmerge_req]

    # Build Header Table 1
    # [AM, 10-15 ngày, Trên 15 ngày, Tồn hoạt động, Đã gán, Chưa gán, Tỉ lệ gán (%), Đã xử lý xong, Tổng đơn, Mốc 7h30, Mốc Hiện tại, +/-]
    fixed_headers = ['AM', 'Tồn 10-15 ngày', 'Tồn Trên 15 ngày', 'Tổng tồn hoạt động', 'Đã gán (Có chuyến)', 'Chưa gán', 'Tỉ lệ gán (%)', 'Đã xử lý xong', 'Tổng đơn (>10 ngày)']
    headers_t1 = list(fixed_headers)
    for i, snap in enumerate(history):
        headers_t1.append(f"Tồn (mốc {snap['time']})")
        if i > 0:
            headers_t1.append("'+/- so với trước")

    grid_values = []
    # Row 1: Title
    grid_values.append(['Đơn aging >10 ngày & Xử lý của AM'] + [''] * (len(headers_t1) - 1))
    requests.append(merge_request(ws_pivot.id, 0, 1, 0, 9))
    requests.append(merge_request(ws_pivot.id, 0, 1, 9, len(headers_t1)))
    requests.append(cell_format_request(ws_pivot.id, 0, 1, 0, len(headers_t1), {
        "textFormat": {"bold": True, "fontSize": 13, "fontFamily": "Arial"},
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "MIDDLE"
    }))
    requests.append(row_height_request(ws_pivot.id, 0, 1, 35))

    # Row 2: Headers
    grid_values.append(headers_t1)
    requests.append(cell_format_request(ws_pivot.id, 1, 2, 0, len(headers_t1), {
        "backgroundColor": make_color("#1e3a8a"), # Navy
        "textFormat": {"bold": True, "fontSize": 10, "fontFamily": "Arial", "foregroundColor": make_color("#FFFFFF")},
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "MIDDLE"
    }))
    requests.append(row_height_request(ws_pivot.id, 1, 2, 32))

    # Rows: AMs
    am_rows_start = 2
    sorted_ams_by_unassigned = sorted(am_names, key=lambda x: pivot_map[x]['active_count'], reverse=True)

    for r_idx, am in enumerate(sorted_ams_by_unassigned):
        row_idx = am_rows_start + r_idx
        stats = pivot_map[am]
        
        # Tỉ lệ gán (%)
        pct_assigned = round(stats['assigned'] / stats['active_count'] * 100) if stats['active_count'] > 0 else 0
        pct_str = f"{pct_assigned}%"
        
        row = [
            am,
            stats['10 - 15 ngày'],
            stats['Trên 15 ngày'],
            stats['active_count'],
            stats['assigned'],
            stats['unassigned'],
            pct_str,
            stats['processed'],
            stats['total_orders']
        ]
        
        # Lịch sử snapshots
        for h_idx, snap in enumerate(history):
            snap_val = snap["totals"].get(am, 0)
            row.append(snap_val)
            if h_idx > 0:
                prev_val = history[h_idx - 1]["totals"].get(am, 0)
                diff = snap_val - prev_val
                row.append("Không đổi" if diff == 0 else f"+ {diff}" if diff > 0 else f"- {abs(diff)}")
                
        grid_values.append(row)
        
        # Định dạng
        bg_color = "#E3F2FD" if r_idx % 2 == 0 else "#FFFFFF"
        # Cột AM
        requests.append(cell_format_request(ws_pivot.id, row_idx, row_idx+1, 0, 1, {
            "backgroundColor": make_color("#E3F2FD"),
            "textFormat": {"bold": True, "fontFamily": "Arial", "fontSize": 9},
            "horizontalAlignment": "LEFT",
            "verticalAlignment": "MIDDLE"
        }))
        # Các cột số liệu
        requests.append(cell_format_request(ws_pivot.id, row_idx, row_idx+1, 1, 9, {
            "backgroundColor": make_color(bg_color),
            "textFormat": {"fontFamily": "Arial", "fontSize": 9},
            "horizontalAlignment": "CENTER",
            "verticalAlignment": "MIDDLE"
        }))
        # Highlight cột Trên 15 ngày
        requests.append(cell_format_request(ws_pivot.id, row_idx, row_idx+1, 2, 3, {
            "backgroundColor": make_color("#FFE4E6"), # Rose light
            "textFormat": {"bold": True, "fontFamily": "Arial", "fontSize": 9, "foregroundColor": make_color("#9F1239")}
        }))
        # Highlight cột Chưa gán
        requests.append(cell_format_request(ws_pivot.id, row_idx, row_idx+1, 5, 6, {
            "backgroundColor": make_color("#FEF3C7"), # Amber light
            "textFormat": {"bold": True, "fontFamily": "Arial", "fontSize": 9, "foregroundColor": make_color("#78350F")}
        }))
        
        # Highlight cột Lịch sử hiện tại
        color_info = SLOT_COLORS[0]
        requests.append(cell_format_request(ws_pivot.id, row_idx, row_idx+1, 9, 10, {
            "backgroundColor": make_color(color_info["data_bg"]),
            "textFormat": {"bold": True, "fontFamily": "Arial", "fontSize": 9},
            "horizontalAlignment": "CENTER",
            "verticalAlignment": "MIDDLE"
        }))
        for idx in range(1, history_len):
            col_offset = 8 + idx * 2
            color_info = SLOT_COLORS[idx % len(SLOT_COLORS)]
            requests.append(cell_format_request(ws_pivot.id, row_idx, row_idx+1, col_offset, col_offset+1, {
                "backgroundColor": make_color(color_info["data_bg"]),
                "textFormat": {"bold": True, "fontFamily": "Arial", "fontSize": 9},
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE"
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

    # AM Total row
    total_row_idx = am_rows_start + len(sorted_ams_by_unassigned)
    tot_active_count = sum(pivot_map[am]['active_count'] for am in am_names)
    tot_assigned = sum(pivot_map[am]['assigned'] for am in am_names)
    tot_pct_assigned = round(tot_assigned / tot_active_count * 100) if tot_active_count > 0 else 0
    
    total_row_val = [
        'TỔNG CỘNG',
        sum(pivot_map[am]['10 - 15 ngày'] for am in am_names),
        sum(pivot_map[am]['Trên 15 ngày'] for am in am_names),
        tot_active_count,
        tot_assigned,
        sum(pivot_map[am]['unassigned'] for am in am_names),
        f"{tot_pct_assigned}%",
        sum(pivot_map[am]['processed'] for am in am_names),
        sum(pivot_map[am]['total_orders'] for am in am_names)
    ]
    
    for h_idx, snap in enumerate(history):
        t_sum = sum(snap["totals"].get(am, 0) for am in am_names)
        total_row_val.append(t_sum)
        if h_idx > 0:
            prev_t_sum = sum(history[h_idx - 1]["totals"].get(am, 0) for am in am_names)
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
    requests.append(cell_format_request(ws_pivot.id, total_row_idx, total_row_idx+1, 9, 10, {
        "backgroundColor": make_color(color_info["total_bg"])
    }))
    for idx in range(1, history_len):
        col_offset = 8 + idx * 2
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
        
    requests.append(row_height_request(ws_pivot.id, total_row_idx, total_row_idx+1, 26))
    requests.append(border_request(ws_pivot.id, 1, total_row_idx + 1, 0, len(headers_t1)))

    grid_values.append([''] * len(headers_t1))

    # Table 2: Top 5 Bưu cục
    bc_start_row_idx = total_row_idx + 2
    grid_values.append(['Top 5 Bưu cục có đơn >10 ngày tồn nhiều nhất'] + [''] * (len(headers_t1) - 1))
    requests.append(merge_request(ws_pivot.id, bc_start_row_idx, bc_start_row_idx+1, 0, 9))
    requests.append(merge_request(ws_pivot.id, bc_start_row_idx, bc_start_row_idx+1, 9, len(headers_t1)))
    requests.append(cell_format_request(ws_pivot.id, bc_start_row_idx, bc_start_row_idx+1, 0, len(headers_t1), {
        "textFormat": {"bold": True, "fontSize": 13, "fontFamily": "Arial"},
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "MIDDLE"
    }))
    requests.append(row_height_request(ws_pivot.id, bc_start_row_idx, bc_start_row_idx+1, 35))

    # Headers BC
    bc_headers = ['Bưu cục', 'Tồn 10-15 ngày', 'Tồn Trên 15 ngày', 'Tổng tồn hoạt động', 'Đã gán (Có chuyến)', 'Chưa gán', 'Tỉ lệ gán (%)', 'Đã xử lý xong', 'Tổng đơn (>10 ngày)']
    for i, snap in enumerate(history):
        bc_headers.append(f"Tồn (mốc {snap['time']})")
        if i > 0:
            bc_headers.append("'+/- so với trước")
            
    grid_values.append(bc_headers)
    requests.append(cell_format_request(ws_pivot.id, bc_start_row_idx+1, bc_start_row_idx+2, 0, len(headers_t1), {
        "backgroundColor": make_color("#1e3a8a"),
        "textFormat": {"bold": True, "fontSize": 10, "fontFamily": "Arial", "foregroundColor": make_color("#FFFFFF")},
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "MIDDLE"
    }))
    requests.append(row_height_request(ws_pivot.id, bc_start_row_idx+1, bc_start_row_idx+2, 32))

    for r_idx, bc_name in enumerate(top5_bc_names):
        row_idx = bc_start_row_idx + 2 + r_idx
        stats = bc_stats[bc_name]
        
        pct_assigned = round(stats['assigned'] / stats['active_count'] * 100) if stats['active_count'] > 0 else 0
        pct_str = f"{pct_assigned}%"
        
        row = [
            bc_name,
            stats['10 - 15 ngày'],
            stats['Trên 15 ngày'],
            stats['active_count'],
            stats['assigned'],
            stats['unassigned'],
            pct_str,
            stats['processed'],
            stats['total_orders']
        ]
        
        for h_idx, snap in enumerate(history):
            snap_val = snap["bcTotals"].get(bc_name, 0)
            row.append(snap_val)
            if h_idx > 0:
                prev_val = history[h_idx - 1]["bcTotals"].get(bc_name, 0)
                diff = snap_val - prev_val
                row.append("Không đổi" if diff == 0 else f"+ {diff}" if diff > 0 else f"- {abs(diff)}")
                
        grid_values.append(row)
        
        bg_color = "#E3F2FD" if r_idx % 2 == 0 else "#FFFFFF"
        requests.append(cell_format_request(ws_pivot.id, row_idx, row_idx+1, 0, 1, {
            "backgroundColor": make_color("#E3F2FD"),
            "textFormat": {"bold": True, "fontFamily": "Arial", "fontSize": 9},
            "horizontalAlignment": "LEFT",
            "verticalAlignment": "MIDDLE"
        }))
        requests.append(cell_format_request(ws_pivot.id, row_idx, row_idx+1, 1, 9, {
            "backgroundColor": make_color(bg_color),
            "textFormat": {"fontFamily": "Arial", "fontSize": 9},
            "horizontalAlignment": "CENTER",
            "verticalAlignment": "MIDDLE"
        }))
        requests.append(cell_format_request(ws_pivot.id, row_idx, row_idx+1, 2, 3, {
            "backgroundColor": make_color("#FFE4E6"),
            "textFormat": {"bold": True, "fontFamily": "Arial", "fontSize": 9, "foregroundColor": make_color("#9F1239")}
        }))
        requests.append(cell_format_request(ws_pivot.id, row_idx, row_idx+1, 5, 6, {
            "backgroundColor": make_color("#FEF3C7"),
            "textFormat": {"bold": True, "fontFamily": "Arial", "fontSize": 9, "foregroundColor": make_color("#78350F")}
        }))
        
        color_info = SLOT_COLORS[0]
        requests.append(cell_format_request(ws_pivot.id, row_idx, row_idx+1, 9, 10, {
            "backgroundColor": make_color(color_info["data_bg"]),
            "textFormat": {"bold": True, "fontFamily": "Arial", "fontSize": 9},
            "horizontalAlignment": "CENTER",
            "verticalAlignment": "MIDDLE"
        }))
        for idx in range(1, history_len):
            col_offset = 8 + idx * 2
            color_info = SLOT_COLORS[idx % len(SLOT_COLORS)]
            requests.append(cell_format_request(ws_pivot.id, row_idx, row_idx+1, col_offset, col_offset+1, {
                "backgroundColor": make_color(color_info["data_bg"]),
                "textFormat": {"bold": True, "fontFamily": "Arial", "fontSize": 9},
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE"
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

    grid_values.append([''] * len(headers_t1))
    grid_values.append([''] * len(headers_t1))

    # Table 3: 8-Day snapshot history
    day_start_row_idx = bc_end_row_idx + 2
    grid_values.append(['Đơn >10 ngày — Mốc 7h30 hằng ngày'] + [''] * 8)
    requests.append(merge_request(ws_pivot.id, day_start_row_idx, day_start_row_idx+1, 0, 9))
    requests.append(cell_format_request(ws_pivot.id, day_start_row_idx, day_start_row_idx+1, 0, 9, {
        "textFormat": {"bold": True, "fontSize": 13, "fontFamily": "Arial"},
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "MIDDLE"
    }))
    requests.append(row_height_request(ws_pivot.id, day_start_row_idx, day_start_row_idx+1, 35))

    # Daily snapshots date list
    anchor_dt = datetime.strptime(today_str, "%Y-%m-%d")
    date_keys = []
    date_labels = []
    for d in range(8):
        dt = anchor_dt - timedelta(days=d)
        date_keys.append(dt.strftime("%Y%m%d"))
        date_labels.append(f"Ngày N-{d}\n({dt.strftime('%d/%m')})" if d > 0 else f"Ngày N\n({dt.strftime('%d/%m')})")

    grid_values.append(['AM'] + date_labels)
    requests.append(cell_format_request(ws_pivot.id, day_start_row_idx+1, day_start_row_idx+2, 0, 9, {
        "backgroundColor": make_color("#1e3a8a"),
        "textFormat": {"bold": True, "fontSize": 10, "fontFamily": "Arial", "foregroundColor": make_color("#FFFFFF")},
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "MIDDLE",
        "wrapStrategy": "WRAP"
    }))
    requests.append(row_height_request(ws_pivot.id, day_start_row_idx+1, day_start_row_idx+2, 40))

    daily_snapshots = state["daily_snapshots"]
    day_am_start_idx = day_start_row_idx + 2
    for r_idx, am in enumerate(sorted_ams_by_unassigned):
        row_idx = day_am_start_idx + r_idx
        row = [am]
        requests.append(cell_format_request(ws_pivot.id, row_idx, row_idx+1, 0, 1, {
            "backgroundColor": make_color("#E3F2FD"),
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
    day_total_row_idx = day_am_start_idx + len(sorted_ams_by_unassigned)
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

    # Căn chỉnh kích thước cột
    col_widths = {0: 300, 1: 100, 2: 100, 3: 120, 4: 130, 5: 100, 6: 100, 7: 110, 8: 120, 9: 120, 10: 120, 11: 120}
    for c_idx, w in col_widths.items():
        requests.append(col_width_request(ws_pivot.id, c_idx, c_idx+1, w))

    # Xử lý grid
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

    # 8. Tách đơn theo AM và cập nhật các sheet AM
    print("📂 Đang tách đơn theo AM...")
    am_groups = {}
    for o in active_over10_orders:
        am = o['am']
        if am:
            if am not in am_groups:
                am_groups[am] = []
            am_groups[am].append(o['row_raw'])
            
    # List all worksheets and index them by their NFC normalized title
    all_worksheets_raw = sh.worksheets()
    
    # We want to keep system sheets. Let's define the set of normalized system sheet titles
    system_sheets = {
        unicodedata.normalize('NFC', name).strip() for name in 
        ["Cơ cấu", "data LM", "PIVOT", "Lượt gán", "PUSH REGION", "Đơn giao aging trên 5 ngày", "Đơn 48h chưa có lần giao nào"]
    }
    
    # Normalize our active AM list
    active_am_set = {unicodedata.normalize('NFC', name).strip() for name in am_names}
    
    # Track which worksheets we are keeping for each active AM
    worksheets_to_keep = {}
    seen_normalized_titles = set()
    
    for ws in all_worksheets_raw:
        norm_title = unicodedata.normalize('NFC', ws.title).strip()
        if norm_title in system_sheets:
            continue
        elif norm_title in active_am_set:
            if norm_title not in seen_normalized_titles:
                seen_normalized_titles.add(norm_title)
                worksheets_to_keep[norm_title] = ws
            else:
                print(f"🗑️ Xóa tab trùng lặp của AM: '{ws.title}' (ID: {ws.id})")
                try:
                    sh.del_worksheet(ws)
                except Exception as e:
                    print(f"⚠️ Không thể xóa tab trùng lặp '{ws.title}': {e}")
        else:
            print(f"🗑️ Xóa tab không hoạt động hoặc không thuộc mốc >10 ngày: '{ws.title}' (ID: {ws.id})")
            try:
                sh.del_worksheet(ws)
            except Exception as e:
                print(f"⚠️ Không thể xóa tab '{ws.title}': {e}")
                
    am_links = {}
    for am_name in am_names:
        am_rows = am_groups.get(am_name, [])
        norm_am_name = unicodedata.normalize('NFC', am_name).strip()
        
        if norm_am_name in worksheets_to_keep:
            ws_am = worksheets_to_keep[norm_am_name]
            if ws_am.title != am_name:
                print(f"🔄 Chuẩn hóa tên tab từ '{ws_am.title}' thành '{am_name}'")
                try:
                    ws_am.update_title(am_name)
                except Exception as e:
                    print(f"⚠️ Cảnh báo: Không thể đổi tên tab: {e}")
            ws_am.clear()
        else:
            print(f"➕ Tạo tab mới cho AM: '{am_name}'")
            ws_am = sh.add_worksheet(title=am_name, rows=str(max(100, len(am_rows) + 50)), cols="15")
        
        ws_am.update([aging_header] + am_rows)
        am_links[am_name] = f"https://docs.google.com/spreadsheets/d/{SHEET_KEY}/edit#gid={ws_am.id}"
        
        sh.batch_update({
            "requests": [
                cell_format_request(ws_am.id, 0, 1, 0, len(aging_header), {
                    "backgroundColor": make_color("#1e3a8a"),
                    "textFormat": {"bold": True, "fontSize": 10, "fontFamily": "Arial", "foregroundColor": make_color("#FFFFFF")},
                    "horizontalAlignment": "CENTER",
                    "verticalAlignment": "MIDDLE"
                }),
                row_height_request(ws_am.id, 0, 1, 28)
            ]
        })
        
    print("✔️ Đã cập nhật xong tất cả các sheet AM.")

    # 9. Render HTML và chụp ảnh PIVOT
    print("📸 Khởi chạy Playwright để chụp bảng Table 1 & Table 2...")
    try:
        generate_and_send_colored_table(pivot_map, am_names, group_labels, history, top5_bc_names, bc_stats, current_time, today_str, top5_bcs, am_links, state)
    except Exception as e:
        print(f"❌ Lỗi khi render/gửi báo cáo: {e}")
        import traceback
        traceback.print_exc()

def generate_and_send_colored_table(pivot_map, am_names, group_labels, history, top5_bc_names, bc_stats, current_time, today_str, top5_bcs, am_links, state):
    extra_headers_html = ""
    for i, snap in enumerate(history):
        color_info = SLOT_COLORS[i % len(SLOT_COLORS)]
        h_bg = color_info["header_bg"]
        extra_headers_html += f'<th style="background: {h_bg} !important; color: #FFFFFF; text-align: center;">Tồn (mốc {snap["time"]})</th>'
        if i > 0:
            extra_headers_html += f'<th style="background: {h_bg} !important; color: #FFFFFF; text-align: center;">+/-</th>'

    def make_delta_badge_html(delta_val):
        if delta_val == 0:
            return '<span class="delta-none">—</span>'
        is_up = delta_val > 0
        num_val = abs(delta_val)
        if is_up:
            return f'<span class="delta-badge delta-red">▲ {num_val:,}</span>'
        else:
            return f'<span class="delta-badge delta-green">▼ {num_val:,}</span>'

    def format_cell_html(val, group_name):
        if val == 0:
            return "<td class='zero-val'>-</td>"
        if group_name == '10 - 15 ngày':
            return f"<td><span class='badge badge-orange'>{val:,}</span></td>"
        elif group_name == 'Trên 15 ngày':
            return f"<td><span class='badge badge-red'>{val:,}</span></td>"
        return f"<td>{val:,}</td>"

    # Table 1 Rows
    t1_rows = ""
    sorted_ams = sorted(am_names, key=lambda x: pivot_map[x]['active_count'], reverse=True)
    for am in sorted_ams:
        stats = pivot_map[am]
        pct_assigned = round(stats['assigned'] / stats['active_count'] * 100) if stats['active_count'] > 0 else 0
        
        t1_rows += f"<tr><td class='left-align bold-text'>{am}</td>"
        t1_rows += format_cell_html(stats['10 - 15 ngày'], '10 - 15 ngày')
        t1_rows += format_cell_html(stats['Trên 15 ngày'], 'Trên 15 ngày')
        
        # Tồn hoạt động, Đã gán, Chưa gán, Tỉ lệ, Đã xử lý, Tổng
        t1_rows += f"<td class='bold-text'>{stats['active_count']:,}</td>"
        t1_rows += f"<td><span class='badge badge-green'>{stats['assigned']:,}</span></td>" if stats['assigned'] > 0 else "<td class='zero-val'>-</td>"
        t1_rows += f"<td><span class='badge badge-amber'>{stats['unassigned']:,}</span></td>" if stats['unassigned'] > 0 else "<td class='zero-val'>-</td>"
        t1_rows += f"<td class='bold-text'>{pct_assigned}%</td>"
        t1_rows += f"<td class='processed-val'>{stats['processed']:,}</td>" if stats['processed'] > 0 else "<td class='zero-val'>-</td>"
        t1_rows += f"<td>{stats['total_orders']:,}</td>"
        
        for h_idx, snap in enumerate(history):
            color_info = SLOT_COLORS[h_idx % len(SLOT_COLORS)]
            d_bg = color_info["data_bg"]
            snap_val = snap["totals"].get(am, 0)
            t1_rows += f"<td class='bold-text' style='background-color: {d_bg} !important;'>{snap_val:,}</td>" if snap_val > 0 else f"<td class='zero-val' style='background-color: {d_bg} !important;'>-</td>"
            if h_idx > 0:
                prev_val = history[h_idx - 1]["totals"].get(am, 0)
                diff = snap_val - prev_val
                badge = make_delta_badge_html(diff)
                td_bg = '#FFCDD2' if diff > 0 else '#C8E6C9' if diff < 0 else d_bg
                t1_rows += f"<td style='background-color: {td_bg} !important;'>{badge}</td>"
        t1_rows += "</tr>"

    # Table 1 Grand Total row
    t1_tot_active = sum(pivot_map[am]['active_count'] for am in am_names)
    t1_tot_assigned = sum(pivot_map[am]['assigned'] for am in am_names)
    t1_tot_pct_assigned = round(t1_tot_assigned / t1_tot_active * 100) if t1_tot_active > 0 else 0
    
    t1_total_row = "<tr class='total-row'><td class='left-align'>TỔNG CỘNG</td>"
    t1_total_row += f"<td>{sum(pivot_map[am]['10 - 15 ngày'] for am in am_names):,}</td>"
    t1_total_row += f"<td>{sum(pivot_map[am]['Trên 15 ngày'] for am in am_names):,}</td>"
    t1_total_row += f"<td>{t1_tot_active:,}</td>"
    t1_total_row += f"<td>{t1_tot_assigned:,}</td>"
    t1_total_row += f"<td>{sum(pivot_map[am]['unassigned'] for am in am_names):,}</td>"
    t1_total_row += f"<td>{t1_tot_pct_assigned}%</td>"
    t1_total_row += f"<td>{sum(pivot_map[am]['processed'] for am in am_names):,}</td>"
    t1_total_row += f"<td>{sum(pivot_map[am]['total_orders'] for am in am_names):,}</td>"
    
    for h_idx, snap in enumerate(history):
        color_info = SLOT_COLORS[h_idx % len(SLOT_COLORS)]
        t_bg = color_info["total_bg"]
        t_sum = sum(snap["totals"].get(am, 0) for am in am_names)
        t1_total_row += f"<td style='background-color: {t_bg} !important; color: #000000; font-weight: 800;'>{t_sum:,}</td>" if t_sum > 0 else f"<td class='zero-val' style='background-color: {t_bg} !important;'>-</td>"
        if h_idx > 0:
            prev_t_sum = sum(history[h_idx - 1]["totals"].get(am, 0) for am in am_names)
            diff = t_sum - prev_t_sum
            badge = make_delta_badge_html(diff)
            td_bg = '#FFCDD2' if diff > 0 else '#C8E6C9' if diff < 0 else t_bg
            t1_total_row += f"<td style='background-color: {td_bg} !important;'>{badge}</td>"
    t1_total_row += "</tr>"

    # Table 2 Rows (BCs)
    t2_rows = ""
    for bc_name in top5_bc_names:
        stats = bc_stats[bc_name]
        pct_assigned = round(stats['assigned'] / stats['active_count'] * 100) if stats['active_count'] > 0 else 0
        
        t2_rows += f"<tr><td class='left-align bold-text'>{bc_name}</td>"
        t2_rows += format_cell_html(stats['10 - 15 ngày'], '10 - 15 ngày')
        t2_rows += format_cell_html(stats['Trên 15 ngày'], 'Trên 15 ngày')
        t2_rows += f"<td class='bold-text'>{stats['active_count']:,}</td>"
        t2_rows += f"<td><span class='badge badge-green'>{stats['assigned']:,}</span></td>" if stats['assigned'] > 0 else "<td class='zero-val'>-</td>"
        t2_rows += f"<td><span class='badge badge-amber'>{stats['unassigned']:,}</span></td>" if stats['unassigned'] > 0 else "<td class='zero-val'>-</td>"
        t2_rows += f"<td class='bold-text'>{pct_assigned}%</td>"
        t2_rows += f"<td class='processed-val'>{stats['processed']:,}</td>" if stats['processed'] > 0 else "<td class='zero-val'>-</td>"
        t2_rows += f"<td>{stats['total_orders']:,}</td>"
        
        for h_idx, snap in enumerate(history):
            color_info = SLOT_COLORS[h_idx % len(SLOT_COLORS)]
            d_bg = color_info["data_bg"]
            snap_val = snap["bcTotals"].get(bc_name, 0)
            t2_rows += f"<td class='bold-text' style='background-color: {d_bg} !important;'>{snap_val:,}</td>" if snap_val > 0 else f"<td class='zero-val' style='background-color: {d_bg} !important;'>-</td>"
            if h_idx > 0:
                prev_val = history[h_idx - 1]["bcTotals"].get(bc_name, 0)
                diff = snap_val - prev_val
                badge = make_delta_badge_html(diff)
                td_bg = '#FFCDD2' if diff > 0 else '#C8E6C9' if diff < 0 else d_bg
                t2_rows += f"<td style='background-color: {td_bg} !important;'>{badge}</td>"
        t2_rows += "</tr>"

    html_content = f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
    body {{
        font-family: 'Inter', sans-serif;
        background: linear-gradient(135deg, #f8fafc 0%, #e2e8f0 100%);
        margin: 0;
        padding: 40px;
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
        max-width: 1750px;
        width: 100%;
        box-sizing: border-box;
    }}
    .header {{
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 24px;
        border-bottom: 2px solid #e2e8f0;
        padding-bottom: 16px;
    }}
    .header h2 {{
        margin: 0;
        font-family: 'Plus Jakarta Sans', sans-serif;
        font-size: 32px;
        background: linear-gradient(90deg, #1e3a8a 0%, #b91c1c 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-weight: 800;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }}
    .header .time {{
        font-size: 15px;
        color: #1e293b;
        font-weight: 700;
        background: #f1f5f9;
        padding: 8px 16px;
        border-radius: 30px;
        border: 1px solid #cbd5e1;
    }}
    .table-section {{
        width: 100%;
        margin-bottom: 40px;
    }}
    .table-title {{
        font-family: 'Plus Jakarta Sans', sans-serif;
        font-size: 24px;
        font-weight: 800;
        color: #1e3a8a;
        margin-top: 10px;
        margin-bottom: 16px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        border-left: 5px solid #b91c1c;
        padding-left: 12px;
        text-align: left;
    }}
    table {{
        width: 100%;
        border-collapse: separate;
        border-spacing: 0;
        text-align: left;
        border-radius: 12px;
        overflow: hidden;
        border: 1px solid #cbd5e1;
    }}
    th {{
        font-family: 'Plus Jakarta Sans', sans-serif;
        background: #1e3a8a;
        color: #ffffff;
        font-weight: 700;
        font-size: 15px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        padding: 14px 10px;
        text-align: center;
        border: none;
        border-bottom: 2px solid #cbd5e1;
     }}
     th.left-align, td.left-align {{
          text-align: left;
          padding-left: 20px;
      }}
      td {{
          padding: 12px 10px;
          font-size: 16px;
          color: #334155;
          border-bottom: 1px solid #e2e8f0;
          font-weight: 600;
          text-align: center;
          background-color: #ffffff;
      }}
      tr:nth-child(even) td {{
          background-color: #f8fafc;
      }}
      tr:hover td {{
          background-color: #f1f5f9;
      }}
      .bold-text {{
          font-weight: 700;
          color: #0f172a;
      }}
      .badge {{
          display: inline-block;
          padding: 4px 10px;
          border-radius: 8px;
          font-weight: 700;
          font-size: 16px;
          text-align: center;
          min-width: 28px;
      }}
      .badge-orange {{
          background-color: #fffbeb;
          color: #d97706;
          border: 1px solid #fde68a;
      }}
      .badge-red {{
          background-color: #fef2f2;
          color: #dc2626;
          border: 1px solid #fecaca;
      }}
      .badge-green {{
          background-color: #f0fdf4;
          color: #15803d;
          border: 1px solid #bbf7d0;
      }}
      .badge-amber {{
          background-color: #fffbeb;
          color: #b45309;
          border: 1px solid #fde68a;
      }}
      .processed-val {{
          color: #16a34a;
          font-weight: 700;
      }}
      .zero-val {{
          color: #94a3b8;
          font-weight: 400;
      }}
      .delta-badge {{
          display: inline-flex;
          align-items: center;
          gap: 4px;
          padding: 4px 8px;
          border-radius: 6px;
          font-size: 14px;
          font-weight: 700;
      }}
      .delta-red {{
          background-color: #fef2f2;
          color: #dc2626;
          border: 1px solid #fecaca;
      }}
      .delta-green {{
          background-color: #f0fdf4;
          color: #16a34a;
          border: 1px solid #bbf7d0;
      }}
      .delta-none {{
          color: #94a3b8;
          font-weight: 500;
      }}
      .total-row td {{
          background: #fff9c4 !important;
          color: #854d0e;
          font-weight: 800;
          border-top: 2px solid #eab308;
          border-bottom: none;
          padding: 16px 10px;
          font-size: 18px;
      }}
</style>
</head>
<body>
<div id="capture-container">
    <div class="header">
        <h2>Báo cáo đơn Aging > 10 ngày</h2>
        <div class="time">Mốc cập nhật: {current_time} ngày {today_str}</div>
    </div>
    
    <div class="table-section">
        <div class="table-title">Thống kê tiến độ theo AM</div>
        <table>
            <thead>
                <tr>
                    <th rowspan="2" class="left-align" style="border-right: 1px solid #cbd5e1;">AM</th>
                    <th colspan="2" style="border-right: 1px solid #cbd5e1; background: #9a3412;">Phân rã tồn</th>
                    <th colspan="4" style="border-right: 1px solid #cbd5e1; background: #0f766e;">Tình trạng gán giao hôm nay</th>
                    <th rowspan="2" style="border-right: 1px solid #cbd5e1; background: #166534;">Đã xử lý xong</th>
                    <th rowspan="2" style="border-right: 1px solid #cbd5e1;">Tổng đơn (>10n)</th>
                    {extra_headers_html}
                </tr>
                <tr>
                    <th style="background: #ea580c; border-right: 1px solid #cbd5e1;">10 - 15 ngày</th>
                    <th style="background: #dc2626; border-right: 1px solid #cbd5e1;">Trên 15 ngày</th>
                    <th style="background: #115e59;">Tồn hoạt động</th>
                    <th style="background: #14b8a6;">Có chuyến đi</th>
                    <th style="background: #d97706;">Chưa gán</th>
                    <th style="background: #0f766e; border-right: 1px solid #cbd5e1;">Tỉ lệ gán (%)</th>
                </tr>
            </thead>
            <tbody>
                {t1_rows}
                {t1_total_row}
            </tbody>
        </table>
    </div>

    <div class="table-section">
        <div class="table-title">Top 5 Bưu Cục tồn >10 ngày nhiều nhất</div>
        <table>
            <thead>
                <tr>
                    <th rowspan="2" class="left-align" style="border-right: 1px solid #cbd5e1;">Bưu cục</th>
                    <th colspan="2" style="border-right: 1px solid #cbd5e1; background: #9a3412;">Phân rã tồn</th>
                    <th colspan="4" style="border-right: 1px solid #cbd5e1; background: #0f766e;">Tình trạng gán giao</th>
                    <th rowspan="2" style="border-right: 1px solid #cbd5e1; background: #166534;">Đã xử lý xong</th>
                    <th rowspan="2" style="border-right: 1px solid #cbd5e1;">Tổng đơn (>10n)</th>
                    {extra_headers_html}
                </tr>
                <tr>
                    <th style="background: #ea580c; border-right: 1px solid #cbd5e1;">10 - 15 ngày</th>
                    <th style="background: #dc2626; border-right: 1px solid #cbd5e1;">Trên 15 ngày</th>
                    <th style="background: #115e59;">Tồn hoạt động</th>
                    <th style="background: #14b8a6;">Có chuyến đi</th>
                    <th style="background: #d97706;">Chưa gán</th>
                    <th style="background: #0f766e; border-right: 1px solid #cbd5e1;">Tỉ lệ gán (%)</th>
                </tr>
            </thead>
            <tbody>
                {t2_rows}
            </tbody>
        </table>
    </div>
</div>
</body>
</html>
"""

    temp_html_path = os.path.join(BASE_DIR, "temp_table_aging_over10.html")
    with open(temp_html_path, "w", encoding="utf-8") as f:
        f.write(html_content)

    output_image_path = os.path.join(BASE_DIR, "table_aging_over10_color.png")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.set_viewport_size({"width": 1800, "height": 1000})
        page.goto(f"file:///{temp_html_path.replace('\\', '/')}")
        page.wait_for_timeout(1000)
        container = page.locator("#capture-container")
        container.screenshot(path=output_image_path)
        browser.close()

    try:
        os.remove(temp_html_path)
    except:
        pass

    # Gửi ảnh và caption lên GTalk
    gtalk_token = "2067164759710552066:RfgqBJY4QtV18udu1wMEfhmoRNI4hgBv"
    gtalk_channel = "2067164759710552066"
    
    file_name = os.path.basename(output_image_path)
    file_size = os.path.getsize(output_image_path)
    
    with open(output_image_path, 'rb') as f:
        file_bytes = f.read()

    # So sánh với mốc cũ
    comparison_totals = {}
    if len(history) > 1:
        comparison_totals = history[0]["totals"]
        comparison_time_label = f"mốc {history[0]['time']} hôm nay"
    else:
        yesterday_key = (datetime.strptime(today_str, "%Y-%m-%d") - timedelta(days=1)).strftime("%Y%m%d")
        yesterday_snap = state.get("daily_snapshots", {}).get(yesterday_key, {})
        comparison_totals = yesterday_snap.get("totals", {})
        comparison_time_label = "mốc sáng qua (N-1)"

    # Tính toán số liệu tổng vùng
    total_region_orders = sum(pivot_map[am]['total_orders'] for am in am_names)
    total_region_processed = sum(pivot_map[am]['processed'] for am in am_names)
    total_region_active = sum(pivot_map[am]['active_count'] for am in am_names)
    total_region_assigned = sum(pivot_map[am]['assigned'] for am in am_names)
    total_region_unassigned = sum(pivot_map[am]['unassigned'] for am in am_names)
    
    pct_region_processed = round(total_region_processed / total_region_orders * 100) if total_region_orders > 0 else 0
    pct_region_assigned = round(total_region_assigned / total_region_active * 100) if total_region_active > 0 else 0
    pct_region_unassigned = round(total_region_unassigned / total_region_active * 100) if total_region_active > 0 else 0

    caption = f"⚠️ <b>[BÁO CÁO ĐƠN AGING >10 NGÀY] Ngày {today_str}</b>\n"
    caption += f"⏱️ <b>Mốc cập nhật:</b> {current_time}\n"
    caption += f"🔗 <b>Xem chi tiết MVĐ theo AM:</b> <a href=\"https://docs.google.com/spreadsheets/d/{SHEET_KEY}/edit#gid=1560830058\"><b>click xem</b></a>\n\n"
    
    caption += f"📊 <b>TIẾN ĐỘ XỬ LÝ :</b>\n"
    caption += f"  - <b>Đã xử lý :</b> {total_region_processed:,}/{total_region_active:,} chiếm {pct_region_processed}%\n"
    caption += f"  - <b>Chưa gán:</b> {total_region_unassigned:,}/{total_region_active:,} (chiếm {pct_region_unassigned}%)\n\n"
    
    caption += f"🚨 <b>TOP AM CHƯA GÁN ĐƠN XỬ LÝ (Cần đẩy gán):</b>\n"
    sorted_ams_by_unassigned = sorted(am_names, key=lambda x: pivot_map[x]['unassigned'], reverse=True)
    show_count = 0
    for am in sorted_ams_by_unassigned:
        stats = pivot_map[am]
        unassigned = stats['unassigned']
        if unassigned == 0 or show_count >= 5:
            continue
        cur_val = stats['active_count']
        pct_assigned = round(stats['assigned'] / cur_val * 100) if cur_val > 0 else 0
        caption += f"  {show_count+1}: AM <b>{am}</b>: <b>{unassigned}</b> đơn chưa gán (Tỉ lệ gán: {pct_assigned}%) | Còn Tồn: <b>{cur_val}</b> đơn\n"
        show_count += 1

    print("📡 Đang gửi ảnh báo cáo đơn Aging >10 ngày sang GTalk group...")
    
    init_payload = {
        "ChannelId": gtalk_channel,
        "FileName": file_name,
        "FileSize": str(file_size),
        "MimeType": "image/png",
        "Metadata": json.dumps({"width": 1800, "height": 1000}),
        "oaToken": gtalk_token
    }
    
    resp_init = requests.post("https://mbff.ghn.vn/api/gtalk/initiate-upload", json=init_payload)
    if resp_init.status_code == 200:
        init_data = resp_init.json()
        if init_data.get("errorCode") == "success":
            presigned_url = init_data["data"]["PresignedURL"]
            upload_id = init_data["data"]["UploadId"]
            
            resp_put = requests.put(presigned_url, data=file_bytes, headers={"Content-Type": "image/png"})
            if resp_put.status_code == 200:
                resp_comp = requests.post("https://mbff.ghn.vn/api/gtalk/complete-upload", json={"oaToken": gtalk_token, "UploadId": upload_id})
                if resp_comp.status_code == 200:
                    comp_data = resp_comp.json()
                    if comp_data.get("errorCode") == "success":
                        file_id = comp_data["data"]["Id"]
                        
                        send_payload = {
                            "channelId": gtalk_channel,
                            "clientMsgId": str(int(os.path.getmtime(output_image_path) * 1000)),
                            "content": {
                                "parseMode": "HTML",
                                "attachment": {
                                    "caption": caption,
                                    "items": [{"image": {"fileId": file_id, "width": 1800, "height": 1000}}]
                                }
                            },
                            "oaToken": gtalk_token
                        }
                        r_send = requests.post("https://mbff.ghn.vn/api/gtalk/send-message", json=send_payload)
                        if r_send.status_code == 200 and r_send.json().get("errorCode") == "success":
                            print("   ✅ Đã gửi báo cáo sang Gtalk group thành công!")
                        else:
                            print(f"   ❌ Gửi tin nhắn GTalk lỗi: {r_send.text}")

def main():
    current_hour = datetime.now().hour
    bypass_time = len(sys.argv) > 1 and sys.argv[1] == "--force"
    if not bypass_time and not (7 <= current_hour <= 22):
        print(f"💤 Ngoài khung giờ hoạt động (7h - 22h). Hiện tại là {datetime.now().strftime('%H:%M:%S')}. Script sẽ dừng.")
        print("💡 Để chạy bất chấp khung giờ này, vui lòng thêm tham số --force khi chạy (Ví dụ: CHAY_BAO_CAO_AGING_OVER10.bat --force)")
        sys.exit(2)

    if "--clear" in sys.argv:
        print("🧹 Đang xóa lịch sử hôm nay...")
        if os.path.exists(SNAPSHOT_FILE):
            try:
                os.remove(SNAPSHOT_FILE)
                print("✔️ Đã xóa lịch sử hôm nay thành công.")
            except Exception as e:
                print(f"❌ Lỗi khi xóa lịch sử: {e}")

    # Bước 1: Tải dữ liệu thô từ GHN
    if "--no-download" in sys.argv:
        print("⚠️ Bỏ qua bước tải dữ liệu từ GHN. Sử dụng dữ liệu hiện có.")
    else:
        try:
            asyncio.run(download_report_data())
        except Exception as e:
            print(f"❌ Lỗi trong quá trình tải dữ liệu GHN: {e}")
            sys.exit(1)

    # Bước 2: Chạy tính toán thống kê và cập nhật sheet
    try:
        run_calculations()
    except Exception as e:
        print(f"❌ Lỗi tính toán gán đơn: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

    print("🎉 HOÀN THÀNH TẤT CẢ CÔNG VIỆC THÀNH CÔNG!")

if __name__ == "__main__":
    main()
