import os
import io
import sys
import json
import requests
import gspread
import time
import unicodedata
from datetime import datetime
from google.oauth2.service_account import Credentials

# Fix encoding for Windows/Task Scheduler/PowerShell
os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except AttributeError:
    pass

# Paths and Constants
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')

# Google Sheets Config (Aging > 10 Days Spreadsheet)
SHEET_KEY = '1vCxSTNgSpO9ETvVRElGyuGc7lnx7LxLRhAB4-lJMHLU'
SHEET_NAME_LM = 'data LM'
SHEET_NAME_AGING = 'Đơn giao aging trên 5 ngày'

# GTalk Config
GTALK_OA_TOKEN = "2067164759710552066:RfgqBJY4QtV18udu1wMEfhmoRNI4hgBv"
GTALK_CHANNEL_ID = "2067164759710552066"

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

def main():
    print(f"⏰ Bắt đầu kiểm tra đơn aging > 10 ngày chưa đi giao lúc: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    # 1. Connect to Google Sheets
    if not os.path.exists(JSON_FILE):
        print(f"❌ Lỗi: Không tìm thấy file credentials.json tại {JSON_FILE}")
        sys.exit(1)
        
    scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    
    try:
        sh = gc_client.open_by_key(SHEET_KEY)
        
        # Read data LM tab
        ws_lm = sh.worksheet(SHEET_NAME_LM)
        lm_data = ws_lm.get_all_values()
        
        # Read Đơn giao aging trên 5 ngày tab
        ws_aging = sh.worksheet(SHEET_NAME_AGING)
        aging_data = ws_aging.get_all_values()
        
        # Read Cơ cấu tab for AM mapping fallback
        ws_cocau = sh.worksheet("Cơ cấu")
        cocau_data = ws_cocau.get_all_values()
    except Exception as e:
        print(f"❌ Lỗi kết nối Google Sheets: {e}")
        sys.exit(1)
        
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

    # 2. Extract backlog status from data LM
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
                
    # 3. Process Đơn giao aging trên 5 ngày
    aging_header = aging_data[0]
    try:
        ag_order_col = aging_header.index("order_code") if "order_code" in aging_header else aging_header.index("mã đơn")
        ag_bc_col = aging_header.index("bc")
        ag_id_bc_col = aging_header.index("id_bc")
        ag_aging_col = aging_header.index("Aging") if "Aging" in aging_header else aging_header.index("aging")
        ag_group_col = aging_header.index("Nhóm BL")
        ag_am_col = aging_header.index("am_name")
        ag_num_col = aging_header.index("num_deliver") if "num_deliver" in aging_header else -1
    except ValueError as e:
        print(f"❌ Lỗi: Không tìm thấy cột bắt buộc trong tab Đơn giao aging trên 5 ngày. Chi tiết: {e}")
        sys.exit(1)
        
    success_keywords = ['đã giao/ chuyển trả thành công', 'đã giao/chuyển trả thành công', 'n/a', 'thành công']
    unattempted_orders_by_am = {}
    total_unattempted_count = 0
    
    for row in aging_data[1:]:
        if len(row) > max(ag_order_col, ag_bc_col, ag_id_bc_col, ag_aging_col, ag_group_col, ag_am_col):
            order_code = row[ag_order_col].strip()
            if not order_code:
                continue
                
            # Check if active
            status = lm_status.get(order_code, '#N/A')
            is_processed = status in ['#N/A', 'n/a'] or any(sk in status.lower() for sk in success_keywords)
            if is_processed:
                continue
                
            # Check if aging > 10 days
            group_val = row[ag_group_col].strip()
            aging_val = row[ag_aging_col].strip()
            group = get_group_over10(aging_val, group_val)
            if not group:
                continue
                
            # Check if num_deliver is 0 or empty
            num_deliver_val = row[ag_num_col].strip() if (ag_num_col != -1 and len(row) > ag_num_col) else '0'
            try:
                num_attempts = int(float(num_deliver_val))
            except ValueError:
                num_attempts = 0
                
            if num_attempts == 0:
                # Match found! Map AM name
                raw_am = row[ag_am_col].strip()
                if not raw_am or raw_am == '#N/A' or raw_am == '':
                    bc_name = row[ag_bc_col].strip()
                    id_bc = row[ag_id_bc_col].strip()
                    am_name = cocau_map.get(id_bc, cocau_map.get(bc_name, "Không xác định"))
                else:
                    am_name = raw_am
                am_name = unicodedata.normalize('NFC', am_name)
                
                if am_name not in unattempted_orders_by_am:
                    unattempted_orders_by_am[am_name] = []
                    
                unattempted_orders_by_am[am_name].append({
                    "code": order_code,
                    "bc": row[ag_bc_col].strip(),
                    "status": status,
                    "aging_days": aging_val
                })
                total_unattempted_count += 1

    # 4. Format GTalk Notification Message
    today_str = datetime.now().strftime('%d/%m/%Y')
    current_time = datetime.now().strftime('%H:%M')
    
    caption = f"🚨 <b>ĐƠN AGING &gt; 10 NGÀY CHƯA ĐI GIAO LẦN NÀO</b>\n"
    caption += f"<b>Mốc cập nhật:</b> {current_time} ngày {today_str}\n"
    caption += f"<b>Tổng số đơn chưa đi giao:</b> <b>{total_unattempted_count}</b> đơn\n"
    caption += f"<b>Xem chi tiết báo cáo:</b> <a href=\"https://docs.google.com/spreadsheets/d/{SHEET_KEY}/edit#gid=1560830058\"><b>Click xem</b></a>\n\n"
    
    if total_unattempted_count > 0:
        # Sort AMs by count descending
        sorted_ams = sorted(unattempted_orders_by_am.items(), key=lambda x: len(x[1]), reverse=True)
        
        for am, orders in sorted_ams:
            caption += f"AM <b>{am}</b>: <b>{len(orders)}</b> đơn\n"
            # Format detailed list of codes
            order_details = []
            for o in orders:
                order_details.append(f"<code>{o['code']}</code>")
                
            caption += f"  • Danh sách: {', '.join(order_details)}\n"
    else:
        caption += f"Tuyệt vời! Hiện tại không có đơn hàng Aging > 10 ngày nào chưa đi giao lần nào (tất cả đều đã được giao thử ít nhất 1 lần)."

    print("=== NỘI DUNG TIN NHẮN ===")
    print(caption)
    
    # 5. Send to GTalk Group
    print("📡 Đang gửi báo cáo cảnh báo sang GTalk...")
    url = "https://mbff.ghn.vn/api/gtalk/send-message"
    client_msg_id = str(int(time.time() * 1000))
    payload = {
        "channelId": GTALK_CHANNEL_ID,
        "clientMsgId": client_msg_id,
        "content": {
            "parseMode": "HTML",
            "text": caption
        },
        "oaToken": GTALK_OA_TOKEN
    }
    headers = {"Content-Type": "application/json"}
    
    try:
        res = requests.post(url, json=payload, headers=headers, timeout=20, verify=False)
        if res.status_code == 200:
            res_data = res.json()
            if res_data.get("errorCode") == "success":
                print("✅ Đã gửi báo cáo sang GTalk group thành công!")
            else:
                print(f"❌ Gửi tin nhắn GTalk lỗi API: {res_data.get('error')}")
        else:
            print(f"❌ Gửi tin nhắn GTalk lỗi HTTP {res.status_code}: {res.text}")
    except Exception as e:
        print(f"❌ Lỗi kết nối gửi GTalk: {e}")

if __name__ == "__main__":
    main()
