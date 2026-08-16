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
from google.oauth2.service_account import Credentials
from playwright.sync_api import sync_playwright
import unicodedata

def normalize_str(s):
    if not s:
        return ""
    return unicodedata.normalize('NFC', str(s).strip())

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
SNAPSHOT_FILE = os.path.join(BASE_DIR, 'snapshot_aging.json')

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
SHEET_KEY = '1WCzgao34cA_SttyB9ytHfE1qKTNl_3iFqDbEfw3lbyU'
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
    default_token = "Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJvcmdDb2RlIjoiZ2huZXhwcmVzcyIsInBhcnRuZXJDb2RlIjoiIiwic2VlZCI6NTY2NzQ5OTE3MTg1Nzk5MDA1LCJzc29JZCI6IjMwNjYwMjEiLCJ1c2VySWQiOiI2NGUyZTJjMjYyY2FkNTVjNmI4NGVlMGQifQ.shzkOdA0wHNa-N7MCqcEiJyA5CIDyU5ACXc-nvLGKWs"
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
        # Dùng danh sách dòng để cập nhật
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

def get_group(col_k_val):
    val = str(col_k_val).strip().lower()
    if any(k in val for k in ['(a)', '(b)', '(c)']):
        return '5 - 8 ngày'
    elif any(k in val for k in ['(d)', '(e)', '(f)', '(g)', '(h)', '(i)', '(j)']):
        return '8 - 15 ngày'
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
            id_bc = normalize_str(r[0])
            bc_name = normalize_str(r[1])
            am_name = normalize_str(r[3])
            if id_bc:
                cocau_map[id_bc] = am_name
            if bc_name:
                cocau_map[bc_name] = am_name

    # 3. Đọc dữ liệu từ sheet gốc "Đơn giao aging trên 5 ngày"
    print("📖 Đọc dữ liệu từ sheet 'Đơn giao aging trên 5 ngày'...")
    ws_aging = sh.worksheet("Đơn giao aging trên 5 ngày")
    aging_data = ws_aging.get_all_values()
    
    if len(aging_data) < 1:
        print("❌ Lỗi: Sheet 'Đơn giao aging trên 5 ngày' trống.")
        sys.exit(1)
        
    aging_header = aging_data[0]
    
    ag_order_idx = aging_header.index("order_code") if "order_code" in aging_header else aging_header.index("mã đơn")
    ag_bc_idx = aging_header.index("bc")
    ag_id_bc_idx = aging_header.index("id_bc")
    ag_group_idx = aging_header.index("Nhóm BL")
    ag_am_idx = aging_header.index("am_name")

    # 4. Phân tích các đơn hàng active (chưa xử lý xong)
    active_orders = []
    for row in aging_data[1:]:
        if len(row) > max(ag_order_idx, ag_bc_idx, ag_id_bc_idx, ag_group_idx, ag_am_idx):
            order_code = row[ag_order_idx].strip()
            if not order_code:
                continue
            
            status = lm_status.get(order_code, '#N/A')
            is_processed = status in ['#N/A', 'đã giao/ chuyển trả thành công', 'đã giao/chuyển trả thành công', 'n/a', 'thành công']
            if is_processed:
                continue
                
            bc_name = normalize_str(row[ag_bc_idx])
            id_bc = normalize_str(row[ag_id_bc_idx])
            group_val = row[ag_group_idx].strip()
            group = get_group(group_val)
            
            am_name = normalize_str(row[ag_am_idx])
            if not am_name or am_name == '#N/A' or am_name == '':
                am_name = cocau_map.get(id_bc, cocau_map.get(bc_name, "Không xác định"))
            am_name = normalize_str(am_name)
                
            active_orders.append({
                'order_code': order_code,
                'bc': bc_name,
                'id_bc': id_bc,
                'am': am_name,
                'group': group,
                'status': status,
                'row_raw': row
            })

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

    # 5. Quản lý Snapshot (Tải trước để có thông tin lịch sử của AM)
    print("📝 Cập nhật lịch sử mốc thời gian...")
    state = {"last_updated_date": "", "history": [], "daily_snapshots": {}}
    if os.path.exists(SNAPSHOT_FILE):
        try:
            with open(SNAPSHOT_FILE, "r", encoding="utf-8") as f:
                state = json.load(f)
        except:
            pass

    # Normalize state history and daily snapshots keys right away
    if "history" in state:
        for snap in state["history"]:
            if "totals" in snap:
                snap["totals"] = {normalize_str(k): v for k, v in snap["totals"].items()}
            if "bcTotals" in snap:
                snap["bcTotals"] = {normalize_str(k): v for k, v in snap["bcTotals"].items()}
    if "daily_snapshots" in state:
        for dk in list(state["daily_snapshots"].keys()):
            if "totals" in state["daily_snapshots"][dk]:
                state["daily_snapshots"][dk]["totals"] = {normalize_str(k): v for k, v in state["daily_snapshots"][dk]["totals"].items()}

    # Tự động quét và khôi phục lịch sử từ sheet PIVOT hiện có (nếu file snapshot bị trống/mới tinh hoặc thiếu ngày)
    try:
        print("🔍 Đang đọc và kiểm tra lịch sử từ sheet PIVOT hiện tại...")
        ws_pivot = None
        for sname in ["PIVOT aging ", "PIVOT aging", "PIVOT", "Pivot"]:
            try:
                ws_pivot = sh.worksheet(sname)
                break
            except Exception:
                pass
        if not ws_pivot:
            ws_pivot = sh.worksheet("PIVOT")
        rows = ws_pivot.get_all_values()
        header_row_idx = -1
        for idx, row in enumerate(rows):
            if len(row) > 0 and "Đơn aging >5 ngày" in row[0] and "hằng ngày" in row[0]:
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
                am_name = normalize_str(row[0])
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
            
            # Tính lại grandTotal
            for dk in state.get("daily_snapshots", {}):
                totals = state["daily_snapshots"][dk]["totals"]
                state["daily_snapshots"][dk]["grandTotal"] = sum(totals.values())
            print("✅ Đã khôi phục dữ liệu lịch sử từ tab PIVOT thành công.")
    except Exception as e:
        print(f"⚠️ Cảnh báo: Không thể khôi phục lịch sử từ PIVOT: {e}")

    # Check new day
    if state.get("last_updated_date") != today_str:
        # Archive yesterday's morning run
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

    # Lấy danh sách tất cả AM từ Cơ cấu và lịch sử để tránh mất AM khi không có đơn trong mốc hiện tại
    am_set = set()
    for r in cocau_data[1:]:
        if len(r) >= 4:
            am = normalize_str(r[3])
            if am and am.lower() not in ['am', 'không xác định', 'chưa phân', '']:
                am_set.add(am)
    
    # Thêm AM từ đơn hàng hiện tại
    for o in active_orders:
        if o['am']:
            am_set.add(normalize_str(o['am']))
            
    # Thêm AM từ lịch sử trong ngày
    for snap in state.get("history", []):
        for am in snap.get("totals", {}).keys():
            if am:
                am_set.add(normalize_str(am))
                
    # Thêm AM từ lịch sử 8 ngày
    for dk in state.get("daily_snapshots", {}):
        for am in state["daily_snapshots"][dk].get("totals", {}).keys():
            if am:
                am_set.add(normalize_str(am))
                
    EXCLUDED_AMS = ["Huỳnh Tấn Hiền", "Nguyễn Tiến Lực", "Nguyễn Minh Hoàng"]
    am_names = sorted([am for am in list(am_set) if am and am not in EXCLUDED_AMS])
    group_labels = ['5 - 8 ngày', '8 - 15 ngày', 'Trên 15 ngày']

    # unassigned counts per AM and group
    pivot_map = {}
    for am in am_names:
        pivot_map[am] = {'5 - 8 ngày': 0, '8 - 15 ngày': 0, 'Trên 15 ngày': 0, 'total': 0}
    for o in active_orders:
        am = o['am']
        group = o['group']
        if group:
            pivot_map[am][group] += 1
            pivot_map[am]['total'] += 1

    # unassigned counts per BC
    bc_stats = {}
    for o in active_orders:
        bc = o['bc']
        group = o['group']
        am = o['am']
        if group:
            if bc not in bc_stats:
                bc_stats[bc] = {'am': am, '5 - 8 ngày': 0, '8 - 15 ngày': 0, 'Trên 15 ngày': 0, 'total': 0}
            bc_stats[bc][group] += 1
            bc_stats[bc]['total'] += 1

    sorted_bcs = sorted(bc_stats.items(), key=lambda x: x[1]['total'], reverse=True)
    top5_bcs = sorted_bcs[:5]
    top5_bc_names = [x[0] for x in top5_bcs]

    current_am_totals = {am: pivot_map[am]['total'] for am in am_names}
    current_bc_totals = {bc: stats['total'] for bc, stats in bc_stats.items()}

    # 5. Cập nhật mốc Snapshot mới
    print("📝 Ghi nhận mốc snapshot mới...")

    current_snap = {
        "time": current_time,
        "totals": current_am_totals,
        "bcTotals": current_bc_totals
    }

    if len(state["history"]) == 0:
        state["history"].append(current_snap)
        state["daily_snapshots"][today_key] = {
            "totals": current_am_totals,
            "grandTotal": sum(current_am_totals.values())
        }
    elif len(state["history"]) == 1:
        state["history"].append(current_snap)
    else:
        state["history"][1] = current_snap

    with open(SNAPSHOT_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)

    history = state["history"]
    history_len = len(history)

    # 6. Ghi dữ liệu và format sheet PIVOT
    print("📊 Cập nhật dữ liệu PIVOT...")
    ws_pivot = None
    for sname in ["PIVOT aging ", "PIVOT aging", "PIVOT", "Pivot"]:
        try:
            ws_pivot = sh.worksheet(sname)
            break
        except Exception:
            pass
    if not ws_pivot:
        ws_pivot = sh.worksheet("PIVOT")
    ws_pivot.clear()

    # Clear all formatting for rows 1-100 first, to clean up any residual colors/borders
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
    # Unmerge all columns A-T, rows 1-100
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

    # Table 1 Header
    headers_t1 = ['AM'] + group_labels
    for i, snap in enumerate(history):
        headers_t1.append(f"Tổng (mốc {snap['time']})")
        if i > 0:
            headers_t1.append("'+/- so với trước")

    grid_values = []
    # Row 1: Title AM
    grid_values.append(['Đơn aging >5 ngày'] + [''] * (len(headers_t1) - 1))
    requests.append(merge_request(ws_pivot.id, 0, 1, 0, 5))
    requests.append(merge_request(ws_pivot.id, 0, 1, 5, len(headers_t1)))
    requests.append(cell_format_request(ws_pivot.id, 0, 1, 0, len(headers_t1), {
        "textFormat": {"bold": True, "fontSize": 13, "fontFamily": "Arial"},
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "MIDDLE"
    }))
    requests.append(row_height_request(ws_pivot.id, 0, 1, 35))

    # Row 2: Headers AM
    grid_values.append(headers_t1)
    requests.append(cell_format_request(ws_pivot.id, 1, 2, 0, len(headers_t1), {
        "backgroundColor": make_color("#1565C0"),
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

    # Row 3 onwards: AM Rows
    am_rows_start = 2
    sorted_ams_by_unassigned = sorted(am_names, key=lambda x: pivot_map[x]['total'], reverse=True)

    for r_idx, am in enumerate(sorted_ams_by_unassigned):
        row_idx = am_rows_start + r_idx
        row = [am]
        for g in group_labels:
            row.append(pivot_map[am][g])
        for h_idx, snap in enumerate(history):
            snap_val = snap["totals"].get(am, 0)
            row.append(snap_val)
            if h_idx > 0:
                prev_val = history[h_idx - 1]["totals"].get(am, 0)
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

    # AM Total row
    total_row_idx = am_rows_start + len(sorted_ams_by_unassigned)
    total_row_val = ['TỔNG']
    for g in group_labels:
        total_row_val.append(sum(pivot_map[am][g] for am in am_names))
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

    grid_values.append([''] * len(headers_t1))

    # Table 2: Top 5 BCs
    bc_start_row_idx = total_row_idx + 2
    grid_values.append(['Top 5 BC tồn nhiều nhất'] + [''] * (len(headers_t1) - 1))
    requests.append(merge_request(ws_pivot.id, bc_start_row_idx, bc_start_row_idx+1, 0, 5))
    requests.append(merge_request(ws_pivot.id, bc_start_row_idx, bc_start_row_idx+1, 5, len(headers_t1)))
    requests.append(cell_format_request(ws_pivot.id, bc_start_row_idx, bc_start_row_idx+1, 0, len(headers_t1), {
        "textFormat": {"bold": True, "fontSize": 13, "fontFamily": "Arial"},
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "MIDDLE"
    }))
    requests.append(row_height_request(ws_pivot.id, bc_start_row_idx, bc_start_row_idx+1, 35))

    bc_headers = ['Bưu cục'] + group_labels
    for i, snap in enumerate(history):
        bc_headers.append(f"Tổng (mốc {snap['time']})")
        if i > 0:
            bc_headers.append("'+/- so với trước")
    grid_values.append(bc_headers)
    requests.append(cell_format_request(ws_pivot.id, bc_start_row_idx+1, bc_start_row_idx+2, 0, len(headers_t1), {
        "backgroundColor": make_color("#1565C0"),
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

    for r_idx, bc_name in enumerate(top5_bc_names):
        row_idx = bc_start_row_idx + 2 + r_idx
        row = [bc_name]
        for g in group_labels:
            row.append(bc_stats[bc_name].get(g, 0))
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

    grid_values.append([''] * len(headers_t1))
    grid_values.append([''] * len(headers_t1))

    # Table 3: 8-Day snapshot history
    day_start_row_idx = bc_end_row_idx + 2
    grid_values.append(['Đơn aging >5 ngày — Mốc 7h30 hằng ngày'] + [''] * 8)
    requests.append(merge_request(ws_pivot.id, day_start_row_idx, day_start_row_idx+1, 0, 9))
    requests.append(cell_format_request(ws_pivot.id, day_start_row_idx, day_start_row_idx+1, 0, 9, {
        "textFormat": {"bold": True, "fontSize": 13, "fontFamily": "Arial"},
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "MIDDLE"
    }))
    requests.append(row_height_request(ws_pivot.id, day_start_row_idx, day_start_row_idx+1, 35))

    # Daily snapshots date list relative to today_str
    anchor_dt = datetime.strptime(today_str, "%Y-%m-%d")
    date_keys = []
    date_labels = []
    for d in range(8):
        dt = anchor_dt - timedelta(days=d)
        date_keys.append(dt.strftime("%Y%m%d"))
        date_labels.append(f"Ngày N-{d}\n({dt.strftime('%d/%m')})" if d > 0 else f"Ngày N\n({dt.strftime('%d/%m')})")

    grid_values.append(['AM'] + date_labels)
    requests.append(cell_format_request(ws_pivot.id, day_start_row_idx+1, day_start_row_idx+2, 0, 9, {
        "backgroundColor": make_color("#1565C0"),
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

    col_widths = {0: 380, 1: 100, 2: 100, 3: 100, 4: 120, 5: 120, 6: 120, 7: 120, 8: 120, 9: 120}
    for c_idx, w in col_widths.items():
        requests.append(col_width_request(ws_pivot.id, c_idx, c_idx+1, w))

    # Clear grid trailing columns/rows
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

    # 7. Tách đơn theo AM và cập nhật các sheet AM
    print("📂 Đang tách đơn theo AM...")
    am_groups = {}
    for o in active_orders:
        am = o['am']
        if am:
            if am not in am_groups:
                am_groups[am] = []
            am_groups[am].append(o['row_raw'])
            
    # Lấy danh sách worksheets hiện tại để tránh gọi API nhiều lần
    all_worksheets = {ws.title: ws for ws in sh.worksheets()}
    
    am_links = {}
    for am_name in am_names:
        am_rows = am_groups.get(am_name, [])
        if am_name in all_worksheets:
            ws_am = all_worksheets[am_name]
            ws_am.clear()
        else:
            ws_am = sh.add_worksheet(title=am_name, rows=str(max(100, len(am_rows) + 50)), cols="15")
        
        ws_am.update([aging_header] + am_rows)
        # Lưu link của tab AM này
        am_links[am_name] = f"https://docs.google.com/spreadsheets/d/{SHEET_KEY}/edit#gid={ws_am.id}"
        
        # Apply header formatting
        sh.batch_update({
            "requests": [
                cell_format_request(ws_am.id, 0, 1, 0, 14, {
                    "backgroundColor": make_color("#1565C0"),
                    "textFormat": {"bold": True, "fontSize": 10, "fontFamily": "Arial", "foregroundColor": make_color("#FFFFFF")},
                    "horizontalAlignment": "CENTER",
                    "verticalAlignment": "MIDDLE"
                }),
                row_height_request(ws_am.id, 0, 1, 28)
            ]
        })
        
    print("✔️ Đã cập nhật xong tất cả các sheet AM.")

    # 8. Render HTML, chụp ảnh PIVOT bằng Playwright và gửi lên Gtalk
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
        extra_headers_html += f'<th style="background: {h_bg} !important; color: #FFFFFF; text-align: center;">Tổng (mốc {snap["time"]})</th>'
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
        if group_name == '5 - 8 ngày':
            return f"<td><span class='badge badge-blue'>{val:,}</span></td>"
        elif group_name == '8 - 15 ngày':
            return f"<td><span class='badge badge-orange'>{val:,}</span></td>"
        elif group_name == 'Trên 15 ngày':
            return f"<td><span class='badge badge-red'>{val:,}</span></td>"
        return f"<td>{val:,}</td>"

    # Table 1 Rows
    t1_rows = ""
    sorted_ams = sorted(am_names, key=lambda x: pivot_map[x]['total'], reverse=True)
    for am in sorted_ams:
        t1_rows += f"<tr><td class='left-align bold-text'>{am}</td>"
        for g in group_labels:
            t1_rows += format_cell_html(pivot_map[am][g], g)
        for h_idx, snap in enumerate(history):
            color_info = SLOT_COLORS[h_idx % len(SLOT_COLORS)]
            d_bg = color_info["data_bg"]
            snap_val = snap["totals"].get(am, 0)
            t1_rows += f"<td class='bold-text' style='background-color: {d_bg} !important;'>{snap_val:,}</td>" if snap_val > 0 else f"<td class='zero-val' style='background-color: {d_bg} !important;'>-</td>"
            if h_idx > 0:
                prev_val = history[h_idx - 1]["totals"].get(am, 0)
                diff = snap_val - prev_val
                if diff > 0:
                    badge = make_delta_badge_html(diff)
                    td_bg = '#FFCDD2'
                elif diff < 0:
                    badge = make_delta_badge_html(diff)
                    td_bg = '#C8E6C9'
                else:
                    badge = make_delta_badge_html(diff)
                    td_bg = d_bg
                t1_rows += f"<td style='background-color: {td_bg} !important;'>{badge}</td>"
        t1_rows += "</tr>"

    # Table 1 Grand Total row
    t1_total_row = "<tr class='total-row'><td class='left-align'>TỔNG CỘNG</td>"
    for g in group_labels:
        g_sum = sum(pivot_map[am][g] for am in am_names)
        t1_total_row += f"<td>{g_sum:,}</td>" if g_sum > 0 else "<td class='zero-val'>-</td>"
    for h_idx, snap in enumerate(history):
        color_info = SLOT_COLORS[h_idx % len(SLOT_COLORS)]
        t_bg = color_info["total_bg"]
        t_sum = sum(snap["totals"].get(am, 0) for am in am_names)
        t1_total_row += f"<td style='background-color: {t_bg} !important; color: #000000; font-weight: 800;'>{t_sum:,}</td>" if t_sum > 0 else f"<td class='zero-val' style='background-color: {t_bg} !important;'>-</td>"
        if h_idx > 0:
            prev_t_sum = sum(history[h_idx - 1]["totals"].get(am, 0) for am in am_names)
            diff = t_sum - prev_t_sum
            if diff > 0:
                badge = make_delta_badge_html(diff)
                td_bg = '#FFCDD2'
            elif diff < 0:
                badge = make_delta_badge_html(diff)
                td_bg = '#C8E6C9'
            else:
                badge = make_delta_badge_html(diff)
                td_bg = t_bg
            t1_total_row += f"<td style='background-color: {td_bg} !important;'>{badge}</td>"
    t1_total_row += "</tr>"

    # Table 2 Rows
    t2_rows = ""
    for bc_name in top5_bc_names:
        t2_rows += f"<tr><td class='left-align bold-text'>{bc_name}</td>"
        for g in group_labels:
            t2_rows += format_cell_html(bc_stats[bc_name].get(g, 0), g)
        for h_idx, snap in enumerate(history):
            color_info = SLOT_COLORS[h_idx % len(SLOT_COLORS)]
            d_bg = color_info["data_bg"]
            snap_val = snap["bcTotals"].get(bc_name, 0)
            t2_rows += f"<td class='bold-text' style='background-color: {d_bg} !important;'>{snap_val:,}</td>" if snap_val > 0 else f"<td class='zero-val' style='background-color: {d_bg} !important;'>-</td>"
            if h_idx > 0:
                prev_val = history[h_idx - 1]["bcTotals"].get(bc_name, 0)
                diff = snap_val - prev_val
                if diff > 0:
                    badge = make_delta_badge_html(diff)
                    td_bg = '#FFCDD2'
                elif diff < 0:
                    badge = make_delta_badge_html(diff)
                    td_bg = '#C8E6C9'
                else:
                    badge = make_delta_badge_html(diff)
                    td_bg = d_bg
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
        background: linear-gradient(135deg, #f1f5f9 0%, #cbd5e1 100%);
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
        max-width: 1650px;
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
    .header h2 {{
        margin: 0;
        font-family: 'Plus Jakarta Sans', sans-serif;
        font-size: 30px;
        background: linear-gradient(90deg, #1e3a8a 0%, #3b82f6 100%);
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
        font-family: 'Plus Jakarta Sans', sans-serif;
        font-size: 24px;
        font-weight: 800;
        color: #1e3a8a;
        margin-top: 10px;
        margin-bottom: 16px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        border-left: 5px solid #2563eb;
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
        font-family: 'Plus Jakarta Sans', sans-serif;
        background: linear-gradient(180deg, #1e40af 0%, #1e3a8a 100%);
        color: #ffffff;
        font-weight: 700;
        font-size: 18px;
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
          font-size: 20px;
          color: #334155;
          border-bottom: 1px solid #f1f5f9;
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
          padding: 4px 12px;
          border-radius: 8px;
          font-weight: 700;
          font-size: 20px;
          text-align: center;
          min-width: 30px;
      }}
      .badge-blue {{
          background-color: #eff6ff;
          color: #2563eb;
          border: 1px solid #bfdbfe;
      }}
      .badge-orange {{
          background-color: #fffbeb;
          color: #d97706;
          border: 1px solid #fde68a;
          box-shadow: 0 0 10px rgba(217, 119, 6, 0.1);
      }}
      .badge-red {{
          background-color: #fef2f2;
          color: #dc2626;
          border: 1px solid #fecaca;
          box-shadow: 0 0 10px rgba(220, 38, 38, 0.15);
      }}
      .zero-val {{
          color: #94a3b8;
          font-weight: 400;
      }}
      .delta-badge {{
          display: inline-flex;
          align-items: center;
          gap: 4px;
          padding: 4px 10px;
          border-radius: 6px;
          font-size: 18px;
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
          background: #fef08a !important;
          color: #854d0e;
          font-weight: 800;
          border-top: 2px solid #eab308;
          border-bottom: none;
          padding: 14px 10px;
          font-size: 22px;
      }}
</style>
</head>
<body>
<div id="capture-container">
    <div class="header">
        <h2>Báo cáo đơn Aging > 5 ngày</h2>
        <div class="time">Mốc cập nhật: {current_time} ngày {today_str}</div>
    </div>
    
    <div class="tables-wrapper">
        <div class="table-section">
            <div class="table-title">Thống kê theo AM</div>
            <table>
                <thead>
                    <tr>
                        <th class="left-align">AM</th>
                        <th>5 - 8 ngày</th>
                        <th style="background: linear-gradient(180deg, #d97706 0%, #b45309 100%); color: #ffffff;">8 - 15 ngày (Aging >8)</th>
                        <th style="background: linear-gradient(180deg, #dc2626 0%, #b91c1c 100%); color: #ffffff;">Trên 15 ngày (Aging >15)</th>
                        {extra_headers_html}
                    </tr>
                </thead>
                <tbody>
                    {t1_rows}
                    {t1_total_row}
                </tbody>
            </table>
        </div>

        <div class="table-section">
            <div class="table-title">Top 5 Bưu Cục tồn nhiều nhất</div>
            <table>
                <thead>
                    <tr>
                        <th class="left-align">Bưu cục</th>
                        <th>5 - 8 ngày</th>
                        <th style="background: linear-gradient(180deg, #d97706 0%, #b45309 100%); color: #ffffff;">8 - 15 ngày (Aging >8)</th>
                        <th style="background: linear-gradient(180deg, #dc2626 0%, #b91c1c 100%); color: #ffffff;">Trên 15 ngày (Aging >15)</th>
                        {extra_headers_html}
                    </tr>
                </thead>
                <tbody>
                    {t2_rows}
                </tbody>
            </table>
        </div>
    </div>
</div>
</body>
</html>
"""

    temp_html_path = os.path.join(BASE_DIR, "temp_table_aging.html")
    with open(temp_html_path, "w", encoding="utf-8") as f:
        f.write(html_content)

    output_image_path = os.path.join(BASE_DIR, "table_aging_color.png")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.set_viewport_size({"width": 1700, "height": 1000})
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
    gtalk_token = "2067164759497973760:RfgqBJY4QtV18udu1wMEfhmoRNI4hgBv"
    gtalk_channel = "2073027751071649792"
    
    file_name = os.path.basename(output_image_path)
    file_size = os.path.getsize(output_image_path)
    
    with open(output_image_path, 'rb') as f:
        file_bytes = f.read()

    # Dynamic GTalk message caption
    comparison_totals = {}
    
    if len(history) > 1:
        comparison_totals = history[0]["totals"]
        comparison_time_label = f"mốc {history[0]['time']} hôm nay"
    else:
        yesterday_key = (datetime.strptime(today_str, "%Y-%m-%d") - timedelta(days=1)).strftime("%Y%m%d")
        yesterday_snap = state.get("daily_snapshots", {}).get(yesterday_key, {})
        comparison_totals = yesterday_snap.get("totals", {})
        comparison_time_label = "mốc sáng qua (N-1)"

    caption = f"📊 <b>Đơn Aging > 5 ngày {today_str}</b>\n"
    caption += f"⏱️ <b>Mốc cập nhật:</b> {current_time}\n"
    caption += f"🔗 <b>Link danh sách đơn cần xử lý của từng AM:</b> <a href=\"https://docs.google.com/spreadsheets/d/{SHEET_KEY}/edit#gid=1662819938\"><b>xem chi tiết</b></a>\n\n"
    
    caption += f"🏆 <b>TOP 5 AM TỒN CAO NGÀY HÔM NAY (so với {comparison_time_label}):</b>\n"
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
                change_text = f" (Giảm {abs(diff)} đơn tồn{pct_str})"
            elif diff > 0:
                change_text = f" (Tồn tăng thêm +{diff} đơn tồn{pct_str})"
            else:
                change_text = " (Không đổi)"
        else:
            change_text = ""
            
        caption += f"  {idx+1}: AM <b>{am}</b>: <b>{current_val}</b> đơn{change_text}\n"

    print("📡 Đang gửi ảnh báo cáo đơn Aging sang GTalk group...")
    
    init_payload = {
        "ChannelId": gtalk_channel,
        "FileName": file_name,
        "FileSize": str(file_size),
        "MimeType": "image/png",
        "Metadata": json.dumps({"width": 1700, "height": 1000}),
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
                                    "items": [{"image": {"fileId": file_id, "width": 1700, "height": 1000}}]
                                }
                            },
                            "oaToken": gtalk_token
                        }
                        r_send = requests.post("https://mbff.ghn.vn/api/gtalk/send-message", json=send_payload)
                        if r_send.status_code == 200 and r_send.json().get("errorCode") == "success":
                            print("   ✅ Đã gửi báo cáo sang Gtalk group thành công!")
                        else:
                            print(f"   ❌ Gửi tin nhắn GTalk lỗi: {r_send.text}")
                            
    # Giữ ảnh cục bộ sau khi gửi
    pass

def main():
    current_hour = datetime.now().hour
    bypass_time = len(sys.argv) > 1 and sys.argv[1] == "--force"
    if not bypass_time and not (7 <= current_hour <= 22):
        print(f"💤 Ngoài khung giờ hoạt động (7h - 22h). Hiện tại là {datetime.now().strftime('%H:%M:%S')}. Script sẽ dừng.")
        print("💡 Để chạy bất chấp khung giờ này, vui lòng thêm tham số --force khi chạy (Ví dụ: CHAY_BAO_CAO_AGING.bat --force)")
        sys.exit(2)

    # Clear snapshot history manually if requested
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
