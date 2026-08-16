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

# Cấu hình của Thủy (Telegram Bot)
API_ID = 33980755
API_HASH = '27cb91d2027884b61393c554a4439dff'
PHONE = '+84368644943'
BOT_NAME = 'ghn_staff_bot'

# Cấu hình Google Sheet
SHEET_KEY = '1l2jZGLFoqxta2jz1RRJDTa1x15nHmXxGa2ZXhWXjM4M'
SHEET_NAME = 'data LM'
TEMP_EXCEL = os.path.join(BASE_DIR, 'temp_ghn.xlsx')
SESSION_FILE = os.path.join(BASE_DIR, 'session_thuy')
SNAPSHOT_FILE = os.path.join(BASE_DIR, 'snapshot_push.json')

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
    trigger_time = datetime.now(timezone.utc) - timedelta(minutes=3)
    
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
        res = requests.get(file_url)
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
    """Đọc dữ liệu, tính toán lượt gán, cập nhật sheet PUSH REGION, Lượt gán và gửi thông báo"""
    print(f"🔄 Bắt đầu phân tích lượt gán đơn lúc: {datetime.now().strftime('%H:%M:%S')}")
    
    # Khởi tạo các biến hiệu số để tránh lỗi scoping / NameError
    diff_unassigned = 0
    diff_assigned = 0
    diff_processed = 0
    diff_active = 0
    
    # Xác thực Google credentials
    gc_client = get_gspread_client(SHEET_KEY)
    sh = gc_client.open_by_key(SHEET_KEY)
    print(f"Spreadsheet mục tiêu: '{sh.title}'")
    
    # 1. Đọc dữ liệu từ sheet "data LM"
    print("📖 Đọc dữ liệu từ tab 'data LM'...")
    ws_lm = sh.worksheet("data LM")
    lm_data = ws_lm.get_all_values()
    
    if len(lm_data) < 2:
        print("❌ Lỗi: Sheet 'data LM' trống hoặc thiếu header.")
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
                
    print(f"   Tìm thấy {len(lm_status)} đơn hàng trong tab 'data LM'.")
    
    # 2. Đọc danh sách đơn từ sheet gốc "Đơn giao aging trên 5 ngày"
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
            
        if "bc" in aging_header:
            ag_bc_idx = aging_header.index("bc")
        else:
            ag_bc_idx = aging_header.index("bc")
            
        if "tinh" in aging_header:
            ag_tinh_idx = aging_header.index("tinh")
        else:
            ag_tinh_idx = aging_header.index("tỉnh")
            
        ag_am_idx = aging_header.index("am_name")
    except ValueError as e:
        print(f"❌ Lỗi: Thiếu các cột bắt buộc trong tab 'Đơn giao aging trên 5 ngày'. Chi tiết: {e}")
        sys.exit(1)
        
    target_orders = []
    for row in aging_data[1:]:
        if len(row) > max(ag_order_idx, ag_bc_idx, ag_tinh_idx, ag_am_idx):
            m_don = row[ag_order_idx].strip()
            bc = row[ag_bc_idx].strip()
            tinh = row[ag_tinh_idx].strip()
            am = row[ag_am_idx].strip()
            if m_don:
                target_orders.append((m_don, bc, tinh, am))
                
    print(f"   Tìm thấy {len(target_orders)} đơn cần theo dõi gán.")
    
    # 3. Lấy dữ liệu cũ trong sheet "PUSH REGION" để giữ lại lịch sử
    print("📖 Kiểm tra tab 'PUSH REGION'...")
    ws_push = None
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
            print(f"   Tìm thấy lịch sử theo dõi của {len(existing_tracking)} đơn hàng.")
    except gspread.exceptions.WorksheetNotFound:
        print("   Tạo mới tab 'PUSH REGION'...")
        ws_push = sh.add_worksheet(title="PUSH REGION", rows="1000", cols="15")
        
    # 4. Tính toán và cập nhật lượt gán
    print("✍️ Đối chiếu và tính toán lượt gán đơn...")
    updated_rows = []
    push_headers = [
        'Mã đơn hàng', 'Tên bưu cục hiện tại', 'Tên tỉnh', 'AM', 
        'Trạng thái hiện tại', 'Có chuyến đi giao', 'Lượt gán', 
        'Trạng thái trước đó', 'Cập nhật lúc'
    ]
    
    now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    new_assignments_count = 0
    
    for m_don, bc, tinh, am in target_orders:
        if m_don in lm_status:
            status_now = lm_status[m_don]
        else:
            status_now = "đã giao/ chuyển trả thành công"
            
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
    
    # 5. Cập nhật thống kê vào sheet "Lượt gán"
    print("📊 Cập nhật thống kê lượt gán...")
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
    
    # 6. Cập nhật tab PIVOT và các tab AM bằng Python trực tiếp (độc lập hoàn toàn, không qua Web App)
    print("📊 [Python] Cập nhật tab PIVOT...")
    try:
        # Lọc ra các đơn active (chưa hoàn thành) để làm PIVOT
        active_orders = []
        ag_group_idx = -1
        for h_name in ["nhóm bl", "nhom bl", "nhóm"]:
            if h_name in aging_header:
                ag_group_idx = aging_header.index(h_name)
                break
        
        for row in aging_data[1:]:
            if len(row) > max(ag_order_idx, ag_bc_idx, ag_tinh_idx, ag_am_idx):
                m_don = row[ag_order_idx].strip()
                if not m_don:
                    continue
                
                status_now = lm_status.get(m_don, "đã giao/ chuyển trả thành công")
                is_processed = status_now in ['đã giao/ chuyển trả thành công', 'đã giao/chuyển trả thành công', '#n/a', 'n/a', 'thành công', '#N/A']
                if is_processed:
                    continue
                
                bc = row[ag_bc_idx].strip()
                am = row[ag_am_idx].strip()
                group_val = row[ag_group_idx].strip() if (ag_group_idx != -1 and len(row) > ag_group_idx) else ""
                group = get_group(group_val)
                
                active_orders.append({
                    'order_code': m_don,
                    'bc': bc,
                    'am': am,
                    'group': group,
                    'row_raw': row
                })
        
        am_names = sorted(list(set([o['am'] for o in active_orders])))
        group_labels = ['5 - 8 ngày', '8 - 15 ngày', 'Trên 15 ngày']

        pivot_map = {}
        for am in am_names:
            pivot_map[am] = {'5 - 8 ngày': 0, '8 - 15 ngày': 0, 'Trên 15 ngày': 0, 'total': 0}
        for o in active_orders:
            am = o['am']
            group = o['group']
            if group:
                pivot_map[am][group] += 1
                pivot_map[am]['total'] += 1

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

        today_str = datetime.now().strftime("%Y-%m-%d")
        today_key = datetime.now().strftime("%Y%m%d")
        current_time = datetime.now().strftime("%H:%M")

        # Quản lý Snapshot
        state = {"last_updated_date": "", "history": [], "daily_snapshots": {}}
        if os.path.exists(SNAPSHOT_FILE):
            try:
                with open(SNAPSHOT_FILE, "r", encoding="utf-8") as f:
                    state = json.load(f)
            except:
                pass

        try:
            print("🔍 Đang đọc và kiểm tra lịch sử từ sheet PIVOT hiện tại...")
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
                    totals_dict = state["daily_snapshots"][dk]["totals"]
                    state["daily_snapshots"][dk]["grandTotal"] = sum(totals_dict.values())
                print("✅ Đã khôi phục dữ liệu lịch sử từ tab PIVOT thành công.")
        except Exception as e:
            print(f"⚠️ Cảnh báo: Không thể khôi phục lịch sử từ PIVOT: {e}")

        if state.get("last_updated_date") != today_str:
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

        # Calculate current summary stats for snapshot comparison
        cur_unassigned = 0
        cur_assigned = 0
        cur_processed = 0
        for row in updated_rows:
            status_now = row[4].strip()
            if status_now == "đã giao/ chuyển trả thành công":
                cur_processed += 1
            elif status_now == "Chưa có chuyến đi trong ngày":
                cur_unassigned += 1
            else:
                cur_assigned += 1
        cur_active = cur_unassigned + cur_assigned

        # Migrate any entries in history that are missing "summary" block (robustness fix)
        for snap in state.get("history", []):
            if "summary" not in snap or not snap["summary"]:
                totals_dict = snap.get("totals", {})
                tot_active = sum(totals_dict.values()) if totals_dict else 0
                cur_total_active = cur_unassigned + cur_assigned
                if cur_total_active > 0:
                    unassigned_ratio = cur_unassigned / cur_total_active
                else:
                    unassigned_ratio = 0.35
                est_unassigned = int(round(tot_active * unassigned_ratio))
                est_assigned = tot_active - est_unassigned
                est_processed = cur_processed
                snap["summary"] = {
                    "grand_unassigned": est_unassigned,
                    "grand_assigned": est_assigned,
                    "grand_processed": est_processed,
                    "grand_active": tot_active
                }

        current_snap = {
            "time": current_time,
            "totals": current_am_totals,
            "bcTotals": current_bc_totals,
            "summary": {
                "grand_unassigned": cur_unassigned,
                "grand_assigned": cur_assigned,
                "grand_processed": cur_processed,
                "grand_active": cur_active
            }
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

        # Calculate diffs compared to the first run of the day
        first_snap = state["history"][0]
        first_summary = first_snap.get("summary", {})
        diff_unassigned = cur_unassigned - first_summary.get("grand_unassigned", cur_unassigned)
        diff_assigned = cur_assigned - first_summary.get("grand_assigned", cur_assigned)
        diff_processed = cur_processed - first_summary.get("grand_processed", cur_processed)
        diff_active = cur_active - first_summary.get("grand_active", cur_active)

        with open(SNAPSHOT_FILE, "w", encoding="utf-8") as f:
            json.dump(state, f, ensure_ascii=False, indent=2)

        history = state["history"]
        history_len = len(history)

        ws_pivot = sh.worksheet("PIVOT")
        ws_pivot.clear()

        clear_format_req = {
            "repeatCell": {
                "range": {"sheetId": ws_pivot.id, "startRowIndex": 0, "endRowIndex": 100, "startColumnIndex": 0, "endColumnIndex": 20},
                "cell": {"userEnteredFormat": {}},
                "fields": "userEnteredFormat"
            }
        }
        unmerge_req = {
            "unmergeCells": {
                "range": {"sheetId": ws_pivot.id, "startRowIndex": 0, "endRowIndex": 100, "startColumnIndex": 0, "endColumnIndex": 20}
            }
        }
        requests = [clear_format_req, unmerge_req]

        headers_t1 = ['AM'] + group_labels
        for i, snap in enumerate(history):
            headers_t1.append(f"Tổng (mốc {snap['time']})")
            if i > 0:
                headers_t1.append("'+/- so với trước")

        grid_values = []
        grid_values.append(['Đơn aging >5 ngày'] + [''] * (len(headers_t1) - 1))
        requests.append(merge_request(ws_pivot.id, 0, 1, 0, 5))
        requests.append(merge_request(ws_pivot.id, 0, 1, 5, len(headers_t1)))
        requests.append(cell_format_request(ws_pivot.id, 0, 1, 0, len(headers_t1), {
            "textFormat": {"bold": True, "fontSize": 13, "fontFamily": "Arial"},
            "horizontalAlignment": "CENTER",
            "verticalAlignment": "MIDDLE"
        }))
        requests.append(row_height_request(ws_pivot.id, 0, 1, 35))

        grid_values.append(headers_t1)
        requests.append(cell_format_request(ws_pivot.id, 1, 2, 0, len(headers_t1), {
            "backgroundColor": make_color("#1565C0"),
            "textFormat": {"bold": True, "fontSize": 10, "fontFamily": "Arial", "foregroundColor": make_color("#FFFFFF")},
            "horizontalAlignment": "CENTER",
            "verticalAlignment": "MIDDLE"
        }))
        requests.append(row_height_request(ws_pivot.id, 1, 2, 30))

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
            requests.append(cell_format_request(ws_pivot.id, row_idx, row_idx+1, 4, 5, {
                "backgroundColor": make_color("#FFF9C4"),
                "textFormat": {"bold": True, "fontFamily": "Arial", "fontSize": 9},
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE",
                "numberFormat": {"type": "NUMBER", "pattern": "#,##0"}
            }))
            if history_len > 1:
                requests.append(cell_format_request(ws_pivot.id, row_idx, row_idx+1, 5, 6, {
                    "backgroundColor": make_color("#FFF9C4"),
                    "textFormat": {"bold": True, "fontFamily": "Arial", "fontSize": 9},
                    "horizontalAlignment": "CENTER",
                    "verticalAlignment": "MIDDLE",
                    "numberFormat": {"type": "NUMBER", "pattern": "#,##0"}
                }))
                diff_val = row[6]
                d_bg, d_fg = ("#FFCDD2", "#B71C1C") if diff_val.startswith('+') else ("#C8E6C9", "#1B5E20") if diff_val.startswith('-') else ("#FFFFFF", "#000000")
                requests.append(cell_format_request(ws_pivot.id, row_idx, row_idx+1, 6, 7, {
                    "backgroundColor": make_color(d_bg),
                    "textFormat": {"bold": True, "fontFamily": "Arial", "fontSize": 9, "foregroundColor": make_color(d_fg)},
                    "horizontalAlignment": "CENTER",
                    "verticalAlignment": "MIDDLE"
                }))
            requests.append(row_height_request(ws_pivot.id, row_idx, row_idx+1, 24))

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
        requests.append(cell_format_request(ws_pivot.id, total_row_idx, total_row_idx+1, 0, 1, {"horizontalAlignment": "LEFT"}))
        requests.append(cell_format_request(ws_pivot.id, total_row_idx, total_row_idx+1, 4, 5, {"backgroundColor": make_color("#F9A825")}))
        if history_len > 1:
            requests.append(cell_format_request(ws_pivot.id, total_row_idx, total_row_idx+1, 5, 6, {"backgroundColor": make_color("#F9A825")}))
            t_diff_val = total_row_val[6]
            td_bg, td_fg = ("#FFCDD2", "#B71C1C") if t_diff_val.startswith('+') else ("#C8E6C9", "#1B5E20") if t_diff_val.startswith('-') else ("#FFFFFF", "#000000")
            requests.append(cell_format_request(ws_pivot.id, total_row_idx, total_row_idx+1, 6, 7, {
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
            requests.append(cell_format_request(ws_pivot.id, row_idx, row_idx+1, 4, 5, {
                "backgroundColor": make_color("#FFF9C4"),
                "textFormat": {"bold": True, "fontFamily": "Arial", "fontSize": 9},
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE",
                "numberFormat": {"type": "NUMBER", "pattern": "#,##0"}
            }))
            if history_len > 1:
                requests.append(cell_format_request(ws_pivot.id, row_idx, row_idx+1, 5, 6, {
                    "backgroundColor": make_color("#FFF9C4"),
                    "textFormat": {"bold": True, "fontFamily": "Arial", "fontSize": 9},
                    "horizontalAlignment": "CENTER",
                    "verticalAlignment": "MIDDLE",
                    "numberFormat": {"type": "NUMBER", "pattern": "#,##0"}
                }))
                diff_val = row[6]
                d_bg, d_fg = ("#FFCDD2", "#B71C1C") if diff_val.startswith('+') else ("#C8E6C9", "#1B5E20") if diff_val.startswith('-') else ("#FFFFFF", "#000000")
                requests.append(cell_format_request(ws_pivot.id, row_idx, row_idx+1, 6, 7, {
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

        max_cols_val = max(len(r) for r in grid_values)
        clean_grid = []
        for r in grid_values:
            if len(r) < max_cols_val:
                r = r + [''] * (max_cols_val - len(r))
            clean_grid.append(r)

        end_col_letter = gspread.utils.rowcol_to_a1(1, len(clean_grid[0])).split("1")[0]
        ws_pivot.update(range_name=f"A1:{end_col_letter}{len(clean_grid)}", values=clean_grid, value_input_option="USER_ENTERED")
        sh.batch_update({"requests": requests})
        print("✅ Đã cập nhật xong sheet 'PIVOT' bằng Python trực tiếp.")

        # Tách đơn theo AM và cập nhật các sheet AM
        print("📂 [Python] Đang tách đơn theo AM...")
        am_groups = {}
        for o in active_orders:
            am = o['am']
            if am:
                if am not in am_groups:
                    am_groups[am] = []
                am_groups[am].append(o['row_raw'])
                
        all_worksheets = {ws.title: ws for ws in sh.worksheets()}
        
        for am_name in am_names:
            am_rows = am_groups.get(am_name, [])
            if am_name in all_worksheets:
                ws_am = all_worksheets[am_name]
                ws_am.clear()
            else:
                ws_am = sh.add_worksheet(title=am_name, rows=str(max(100, len(am_rows) + 50)), cols="15")
            
            ws_am.update([aging_data[0]] + am_rows)
            
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
        print("✅ Đã cập nhật xong các sheet AM bằng Python trực tiếp.")
    except Exception as ex_pivot:
        print(f"⚠️ Lỗi cập nhật PIVOT / AM sheets bằng Python: {ex_pivot}")
        import traceback
        traceback.print_exc()
        
    # 7. Vẽ bảng màu và phát sóng lên Gtalk/Telegram
    print("📸 Tạo ảnh bảng màu theo AM và gửi tin nhắn...")
    diff_data = {
        "diff_unassigned": diff_unassigned,
        "diff_assigned": diff_assigned,
        "diff_processed": diff_processed,
        "diff_active": diff_active
    }
    try:
        generate_and_send_colored_table(updated_rows, diff_data)
    except Exception as e:
        print(f"⚠️ Lỗi khi vẽ hoặc gửi bảng màu: {e}")

def generate_and_send_colored_table(updated_rows, diff_data=None):
    import json
    from playwright.sync_api import sync_playwright
    
    if diff_data is None:
        diff_data = {
            "diff_unassigned": 0,
            "diff_assigned": 0,
            "diff_processed": 0,
            "diff_active": 0
        }
        
    diff_unassigned = diff_data.get("diff_unassigned", 0)
    diff_assigned = diff_data.get("diff_assigned", 0)
    diff_processed = diff_data.get("diff_processed", 0)
    diff_active = diff_data.get("diff_active", 0)
    
    stats = {}
    for row in updated_rows:
        am = row[3].strip()
        status_now = row[4].strip()
        
        if not am:
            am = "Không xác định"
            
        if am not in stats:
            stats[am] = {"unassigned": 0, "assigned": 0, "processed": 0, "total": 0}
            
        stats[am]["total"] += 1
        if status_now == "đã giao/ chuyển trả thành công":
            stats[am]["processed"] += 1
        elif status_now == "Chưa có chuyến đi trong ngày":
            stats[am]["unassigned"] += 1
        else:
            stats[am]["assigned"] += 1
            
    sorted_ams = sorted(stats.keys(), key=lambda k: stats[k]["unassigned"], reverse=True)
    
    grand_unassigned = sum(stats[am]["unassigned"] for am in stats)
    grand_assigned = sum(stats[am]["assigned"] for am in stats)
    grand_processed = sum(stats[am]["processed"] for am in stats)
    grand_total = sum(stats[am]["total"] for am in stats)
    grand_active = grand_unassigned + grand_assigned
    grand_pct = (grand_assigned / grand_active * 100) if grand_active > 0 else 0.0
    
    now_dt = datetime.now()
    date_str = now_dt.strftime('%d-%m-%Y')
    time_str = now_dt.strftime('%H:%M')
    
    def format_diff_text(diff, type_col):
        if diff == 0:
            return " (—)"
        arrow = "▲" if diff > 0 else "▼"
        sign = "+" if diff > 0 else ""
        return f" ({arrow}{sign}{abs(diff)})"

    caption_text = (
        f"<b>BÁO CÁO CẬP NHẬT LƯỢT GÁN (AGING &gt;5 NGÀY)</b>\n"
        f"📅 Ngày: {date_str} | ⏱️ Mốc cập nhật: {time_str}\n"
        f"========================\n"
        f"📦 Tổng đơn aging &gt; 5 cần xử lý : {grand_active}{format_diff_text(diff_active, 'active')}\n"
        f"✅ Đã gán chuyến đi: {grand_assigned} đơn{format_diff_text(diff_assigned, 'assigned')}\n"
        f"❌ Chưa gán chuyến đi: {grand_unassigned} đơn{format_diff_text(diff_unassigned, 'unassigned')}\n"
        f"🎉 Đã xử lý (GTC/chuyển trả TC): {grand_processed} đơn{format_diff_text(diff_processed, 'processed')}\n"
        f"\nAnh/chị AM tiếp tục check và gán đơn xử lý nhé\n"
        f"https://docs.google.com/spreadsheets/d/1l2jZGLFoqxta2jz1RRJDTa1x15nHmXxGa2ZXhWXjM4M/edit?gid=1040733966#gid=1040733966"
    )
    
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
        max-width: 650px;
        width: 100%;
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
        font-size: 26px;
        color: #0f172a;
        font-weight: 800;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }}
    table {{
        width: 100%;
        border-collapse: collapse;
        text-align: left;
    }}
    th {{
        font-family: 'Plus Jakarta Sans', sans-serif;
        background-color: #f8fafc;
        color: #475569;
        font-weight: 700;
        font-size: 16px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        padding: 16px 12px;
        border-bottom: 2px solid #e2e8f0;
    }}
    td {{
        padding: 14px 12px;
        font-size: 18px;
        color: #334155;
        border-bottom: 1px solid #f1f5f9;
        font-weight: 500;
    }}
    tr:hover td {{
        background-color: #f8fafc;
    }}
    .am-name {{
        font-weight: 600;
        color: #0f172a;
    }}
    .number {{
        text-align: center;
    }}
    .unassigned-col {{
        color: #e11d48;
        font-weight: 700;
    }}
    .assigned-col {{
        color: #2563eb;
        font-weight: 700;
    }}
    .processed-col {{
        color: #64748b;
        font-weight: 500;
    }}
    .total-col {{
        font-weight: 700;
        color: #0f172a;
    }}
    .rate-badge {{
        display: inline-flex;
        align-items: center;
        gap: 6px;
        padding: 6px 14px;
        border-radius: 9999px;
        font-size: 15px;
        font-weight: 700;
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
    .dot {{
        width: 8px;
        height: 8px;
        border-radius: 50%;
        display: inline-block;
    }}
    .dot-green {{ background-color: #15803d; }}
    .dot-yellow {{ background-color: #a16207; }}
    .dot-red {{ background-color: #b91c1c; }}
    
    .total-row td {{
        background-color: #f8fafc;
        font-weight: 800;
        color: #0f172a;
        border-top: 2px solid #cbd5e1;
        border-bottom: none;
        padding: 18px 12px;
        font-size: 18px;
    }}
</style>
</head>
<body>
<div id="table-container">
    <div class="header">
        <h2>Báo cáo đơn tồn Aging > 5 ngày</h2>
    </div>
    <table>
        <thead>
            <tr>
                <th>AM</th>
                <th style="text-align:center;">Chưa gán</th>
                <th style="text-align:center;">Đã gán</th>
                <th style="text-align:center;">Đã xử lý</th>
                <th style="text-align:center;">Tổng đơn</th>
                <th style="text-align:center;">Tỷ lệ gán</th>
            </tr>
        </thead>
        <tbody>
"""

    for am in sorted_ams:
        s = stats[am]
        unassigned = s["unassigned"]
        assigned = s["assigned"]
        processed = s["processed"]
        total = s["total"]
        active = unassigned + assigned
        rate = (assigned / active * 100) if active > 0 else 0.0
        
        if rate >= 90.0:
            badge_class = "rate-green"
            dot_class = "dot-green"
        elif rate >= 70.0:
            badge_class = "rate-yellow"
            dot_class = "dot-yellow"
        else:
            badge_class = "rate-red"
            dot_class = "dot-red"
            
        html_content += f"""
            <tr>
                <td class="am-name">{am}</td>
                <td class="number unassigned-col">{unassigned}</td>
                <td class="number assigned-col">{assigned}</td>
                <td class="number processed-col">{processed}</td>
                <td class="number total-col">{total}</td>
                <td style="text-align:center;">
                    <span class="rate-badge {badge_class}"><span class="dot {dot_class}"></span>{rate:.1f}%</span>
                </td>
            </tr>
        """

    def format_diff_html(diff, type_col):
        if diff == 0:
            return "<span style='font-size: 14px; color: #64748b; font-weight: normal; margin-left: 4px;'> (—)</span>"
        is_good = False
        if type_col in ['assigned', 'processed']:
            is_good = diff > 0
        elif type_col in ['unassigned', 'active']:
            is_good = diff < 0
        color = "#16a34a" if is_good else "#dc2626"
        sign = "+" if diff > 0 else ""
        arrow = "▲" if diff > 0 else "▼"
        return f"<span style='font-size: 14px; color: {color}; font-weight: bold; margin-left: 4px;'> ({arrow}{sign}{abs(diff)})</span>"

    grand_badge = "rate-yellow" if grand_pct >= 70 else "rate-red"
    grand_dot = "dot-yellow" if grand_pct >= 70 else "dot-red"
    html_content += f"""
            <tr class="total-row">
                <td style="font-family: 'Plus Jakarta Sans', sans-serif; font-weight: 800;">TỔNG CỘNG</td>
                <td class="number unassigned-col">{grand_unassigned}{format_diff_html(diff_unassigned, 'unassigned')}</td>
                <td class="number assigned-col">{grand_assigned}{format_diff_html(diff_assigned, 'assigned')}</td>
                <td class="number processed-col">{grand_processed}{format_diff_html(diff_processed, 'processed')}</td>
                <td class="number total-col">{grand_total}<span style="font-size: 14px; color: #64748b; font-weight: normal; margin-left: 4px;"> (—)</span></td>
                <td style="text-align:center;">
                    <span class="rate-badge {grand_badge}"><span class="dot {grand_dot}"></span>{grand_pct:.1f}%</span>
                </td>
            </tr>
        </tbody>
    </table>
</div>
</body>
</html>
"""

    temp_html_path = os.path.join(BASE_DIR, "table_temp_auto.html")
    with open(temp_html_path, "w", encoding="utf-8") as f:
        f.write(html_content)
        
    output_image_path = os.path.join(BASE_DIR, "table_am_color_auto.png")
    
    print("   Renderer: Khởi chạy Playwright để chụp bảng...")
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
        
    print(f"   Ảnh đã lưu tại: {output_image_path}")
    
    # Gửi ảnh lên GTalk
    gtalk_token = "2067164759710552066:RfgqBJY4QtV18udu1wMEfhmoRNI4hgBv"
    gtalk_channel = "2067164759710552066"
    
    print("   Gửi ảnh bảng màu sang Gtalk group...")
    file_name = os.path.basename(output_image_path)
    file_size = os.path.getsize(output_image_path)
    
    with open(output_image_path, 'rb') as f:
        file_bytes = f.read()
        
    init_payload = {
        "ChannelId": gtalk_channel,
        "FileName": file_name,
        "FileSize": str(file_size),
        "MimeType": "image/png",
        "Metadata": json.dumps({"width": 1200, "height": 800}),
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
                                    "caption": caption_text,
                                    "items": [{"image": {"fileId": file_id, "width": 1200, "height": 800}}]
                                }
                            },
                            "oaToken": gtalk_token
                        }
                        requests.post("https://mbff.ghn.vn/api/gtalk/send-message", json=send_payload)
                        print("   ✅ Đã gửi bảng màu sang Gtalk group thành công!")
                        
    # Gửi ảnh lên Telegram (nếu có bot hoạt động)
    tele_token = "8570130113:AAGXRiUaKBknVpgtm1_i9ZA47JRjAXmB21M"
    tele_chat = "-5058464865"
    print("   Gửi ảnh bảng màu sang Telegram group...")
    tele_url = f"https://api.telegram.org/bot{tele_token}/sendPhoto"
    try:
        with open(output_image_path, 'rb') as f:
            resp_tele = requests.post(tele_url, data={
                "chat_id": tele_chat,
                "caption": caption_text,
                "parse_mode": "HTML"
            }, files={"photo": f}, timeout=15)
            if resp_tele.status_code == 200:
                print("   ✅ Đã gửi bảng màu sang Telegram group thành công!")
            else:
                print(f"   ⚠️ Telegram phản hồi lỗi: {resp_tele.status_code} - {resp_tele.text}")
    except Exception as e:
        print(f"   ⚠️ Lỗi khi gửi Telegram: {e}")

def main():
    current_hour = datetime.now().hour
    bypass_time = len(sys.argv) > 1 and sys.argv[1] == "--force"
    if not bypass_time and not (7 <= current_hour <= 18):
        print(f"💤 Ngoài khung giờ hoạt động (7h - 18h). Hiện tại là {datetime.now().strftime('%H:%M:%S')}. Script sẽ dừng.")
        print("💡 Để chạy bất chấp khung giờ này, vui lòng thêm tham số --force khi chạy (Ví dụ: CHAY_BAO_CAO_PUSH.bat --force)")
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
            import traceback
            traceback.print_exc()
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
