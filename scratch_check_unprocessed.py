import os
import io
import sys
import json
import requests
import gspread
import time
from google.oauth2.service_account import Credentials
from datetime import datetime

# Fix encoding cho Windows/Task Scheduler/PowerShell
os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except AttributeError:
    pass

# Cấu hình đường dẫn
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
HISTORY_FILE = os.path.join(BASE_DIR, 'unprocessed_history.json')

# Cấu hình Google Sheet
SHEET_KEY_DATA = '1l2jZGLFoqxta2jz1RRJDTa1x15nHmXxGa2ZXhWXjM4M'  # File chứa PUSH REGION thô
SHEET_KEY_MAIN = '1WCzgao34cA_SttyB9ytHfE1qKTNl_3iFqDbEfw3lbyU'  # File chính chia sẻ cho AM
scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

# Cấu hình GTalk
GTALK_OA_TOKEN = "2067164759710552066:RfgqBJY4QtV18udu1wMEfhmoRNI4hgBv"
GTALK_CHANNEL_ID = "2067164759710552066"

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

def check_unprocessed_orders():
    print("=== KIỂM TRA ĐƠN HÀNG BỊ DỒN VÀ KHÔNG XỬ LÝ ===")
    
    # 1. Kết nối Google Sheets
    if not os.path.exists(JSON_FILE):
        print(f"❌ Không tìm thấy file credentials.json tại {JSON_FILE}")
        return
        
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    
    sh_data = gc_client.open_by_key(SHEET_KEY_DATA)
    sh_main = gc_client.open_by_key(SHEET_KEY_MAIN)
    
    # 2. Đọc dữ liệu từ tab 'PUSH REGION'
    print("📖 Đọc dữ liệu từ tab 'PUSH REGION'...")
    try:
        ws_push = sh_data.worksheet("PUSH REGION")
        push_data = ws_push.get_all_values()
    except Exception as e:
        print(f"❌ Không đọc được tab 'PUSH REGION'. Lỗi: {e}")
        return

    if len(push_data) < 2:
        print("⚠️ Tab 'PUSH REGION' trống hoặc chỉ có tiêu đề.")
        return

    headers = [h.strip() for h in push_data[0]]
    rows = push_data[1:]
    
    try:
        order_idx = headers.index("Mã đơn hàng")
        bc_idx = headers.index("Tên bưu cục hiện tại")
        tinh_idx = headers.index("Tên tỉnh")
        am_idx = headers.index("AM")
        status_idx = headers.index("Trạng thái hiện tại")
        trip_idx = headers.index("Có chuyến đi giao")
        count_idx = headers.index("Lượt gán")
        time_idx = headers.index("Cập nhật lúc")
    except ValueError as e:
        print(f"❌ Thiếu cột bắt buộc trong sheet. Lỗi: {e}")
        print(f"Các cột đang có: {headers}")
        return

    # 3. Phân loại đơn hàng theo AM
    total_orders = len(rows)
    unassigned_orders = []    # Đơn chưa từng được đi giao (Lượt gán = 0 hoặc Có chuyến = Không)
    failed_delivery_orders = [] # Đơn có đi giao trong ngày nhưng vẫn bị dồn (Chưa giao thành công)
    processed_success_orders = [] # Đơn đã giao/trả thành công
    
    # Gom thống kê theo AM
    am_stats = {}
    
    for row in rows:
        if len(row) <= max(order_idx, status_idx, trip_idx, count_idx):
            continue
            
        m_don = row[order_idx].strip()
        bc = row[bc_idx].strip()
        tinh = row[tinh_idx].strip()
        am = row[am_idx].strip()
        status = row[status_idx].strip().lower()
        trip = row[trip_idx].strip()
        time_str = row[time_idx].strip()
        
        if not am or am == "#N/A":
            am = "Không xác định"
            
        if am not in am_stats:
            am_stats[am] = {"total": 0, "unassigned": 0, "failed": 0, "success": 0}
            
        try:
            luot_gan = int(row[count_idx].strip())
        except ValueError:
            luot_gan = 0
            
        order_info = {
            "mã đơn": m_don,
            "bưu cục": bc,
            "tỉnh": tinh,
            "AM": am,
            "trạng thái": row[status_idx].strip(),
            "lượt gán": luot_gan,
            "có chuyến đi": trip,
            "cập nhật lúc": time_str
        }
        
        # Kiểm tra trạng thái hoàn thành
        is_success = "thành công" in status or "đã giao" in status or "chuyển trả" in status
        
        if is_success:
            processed_success_orders.append(order_info)
            am_stats[am]["success"] += 1
        elif luot_gan == 0 or trip == "Không":
            unassigned_orders.append(order_info)
            am_stats[am]["unassigned"] += 1
            am_stats[am]["total"] += 1
        else:
            failed_delivery_orders.append(order_info)
            am_stats[am]["failed"] += 1
            am_stats[am]["total"] += 1

    # 4. Quản lý lịch sử Số ngày bị dồn liên tiếp không gán
    print("📝 Cập nhật lịch sử lũy kế số ngày bị dồn không gán...")
    history = {"last_processed_date": "", "orders": {}}
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                history = json.load(f)
        except Exception as e:
            print(f"⚠️ Cảnh báo: Không đọc được file lịch sử {HISTORY_FILE}. Tạo mới. Lỗi: {e}")

    today_str = datetime.now().strftime("%Y-%m-%d")
    # Kiểm tra xem đây có phải là ngày mới so với lần chạy trước không
    is_new_day = (history.get("last_processed_date") != today_str)
    
    current_unprocessed = {o["mã đơn"]: o for o in unassigned_orders}
    
    # Cập nhật số ngày bị dồn cho các đơn hàng chưa gán hôm nay
    updated_orders_history = {}
    for m_don, o in current_unprocessed.items():
        if m_don in history.get("orders", {}):
            days = history["orders"][m_don]["consecutive_unassigned_days"]
            # Nếu sang ngày mới thì tăng số ngày bị dồn lên +1
            if is_new_day:
                days += 1
            updated_orders_history[m_don] = {
                "consecutive_unassigned_days": days,
                "last_seen_date": today_str
            }
        else:
            # Đơn mới xuất hiện lần đầu trong danh sách chưa gán
            updated_orders_history[m_don] = {
                "consecutive_unassigned_days": 1,
                "last_seen_date": today_str
            }
            
    # Lưu lại lịch sử mới
    history["last_processed_date"] = today_str
    history["orders"] = updated_orders_history
    
    try:
        with open(HISTORY_FILE, "w", encoding="utf-8") as f:
            json.dump(history, f, ensure_ascii=False, indent=2)
        print("💾 Đã lưu lịch sử lũy kế số ngày dồn.")
    except Exception as e:
        print(f"❌ Không lưu được file lịch sử: {e}")

    # Tạo các dòng chi tiết để ghi vào GSheet
    unprocessed_details_rows = []
    
    # Khởi tạo khóa thống kê ngâm đơn cho các AM
    for am in am_stats:
        am_stats[am]["unassigned_2days_plus"] = 0
        
    for o in unassigned_orders:
        m_don = o["mã đơn"]
        am_name = o["AM"]
        days = history["orders"].get(m_don, {}).get("consecutive_unassigned_days", 1)
        unprocessed_details_rows.append([
            m_don, o["bưu cục"], o["tỉnh"], am_name, o["trạng thái"], 
            o["có chuyến đi"], str(o["lượt gán"]), days, o["cập nhật lúc"]
        ])
        
        # Thống kê số đơn chưa gán liên tiếp từ 2 ngày trở lên
        if days >= 2:
            if am_name not in am_stats:
                am_stats[am_name] = {"total": 0, "unassigned": 0, "failed": 0, "success": 0, "unassigned_2days_plus": 0}
            if "unassigned_2days_plus" not in am_stats[am_name]:
                am_stats[am_name]["unassigned_2days_plus"] = 0
            am_stats[am_name]["unassigned_2days_plus"] += 1

    # Sắp xếp danh sách đơn ghi sheet theo Số ngày dồn giảm dần, rồi đến AM
    unprocessed_details_rows.sort(key=lambda x: (-x[7], x[3]))

    # 5. Ghi danh sách mã đơn KHÔNG XỬ LÝ (Nhóm 1) sang sheet chính (file 1WCz)
    print("✍️ Cập nhật danh sách chi tiết đơn Không Xử Lý sang Google Sheet chính...")
    tab_name = "Đơn Chưa Xử Lý Aging"
    try:
        all_main_worksheets = {ws.title: ws for ws in sh_main.worksheets()}
        if tab_name in all_main_worksheets:
            ws_unprocessed = all_main_worksheets[tab_name]
            ws_unprocessed.clear()
        else:
            ws_unprocessed = sh_main.add_worksheet(title=tab_name, rows=str(max(100, len(unprocessed_details_rows) + 50)), cols="10")
            
        # Ghi tiêu đề và dữ liệu
        unprocessed_headers = [
            'Mã đơn hàng', 'Tên bưu cục hiện tại', 'Tên tỉnh', 'AM', 
            'Trạng thái hiện tại', 'Có chuyến đi giao', 'Lượt gán', 
            'Số ngày chưa gán liên tiếp', 'Cập nhật lúc'
        ]
        ws_unprocessed.update([unprocessed_headers] + unprocessed_details_rows)
        
        # Định dạng Header của tab
        sh_main.batch_update({
            "requests": [
                cell_format_request(ws_unprocessed.id, 0, 1, 0, len(unprocessed_headers), {
                    "backgroundColor": make_color("#E11D48"), # Màu hồng đỏ nổi bật cảnh báo
                    "textFormat": {"bold": True, "fontSize": 10, "fontFamily": "Arial", "foregroundColor": make_color("#FFFFFF")},
                    "horizontalAlignment": "CENTER",
                    "verticalAlignment": "MIDDLE"
                }),
                row_height_request(ws_unprocessed.id, 0, 1, 28)
            ]
        })
        sheet_link = f"https://docs.google.com/spreadsheets/d/{SHEET_KEY_MAIN}/edit#gid={ws_unprocessed.id}"
        print(f"✅ Đã ghi {len(unprocessed_details_rows)} đơn không xử lý vào tab '{tab_name}'!")
    except Exception as e:
        print(f"⚠️ Lỗi ghi danh sách chi tiết đơn chưa xử lý: {e}")
        sheet_link = f"https://docs.google.com/spreadsheets/d/{SHEET_KEY_MAIN}/edit?gid=1040733966#gid=1040733966"

    # 6. Tạo nội dung báo cáo gửi GTalk
    now_str = datetime.now().strftime('%d/%m/%Y %H:%M')
    
    caption = f"🚨 <b>CẢNH BÁO ĐƠN AGING BỊ DỒN - CHƯA GÁN CHUYẾN ĐI GIAO</b>\n"
    caption += f"⏱️ <b>Mốc cập nhật:</b> {now_str}\n"
    caption += f"🔗 <b>Link danh sách đơn CHƯA GÁN:</b> <a href=\"{sheet_link}\"><b>Xem chi tiết tại đây</b></a>\n\n"
    
    # Tính tổng số đơn chưa gán >= 2 ngày
    unassigned_2days_plus_grand = sum(stats.get("unassigned_2days_plus", 0) for stats in am_stats.values())
    
    caption += f"📦 <b>TỔNG HỢP TÌNH TRẠNG (AGING &gt; 5 NGÀY):</b>\n"
    if unassigned_2days_plus_grand > 0:
        caption += f"  • 🚫 <b>Chưa gán giao (Bị dồn):</b> <b>{len(unassigned_orders)}</b> đơn (Trong đó chưa gán liên tiếp &gt;=2 ngày: <b>{unassigned_2days_plus_grand}</b> đơn)\n"
    else:
        caption += f"  • 🚫 <b>Chưa gán giao (Bị dồn):</b> <b>{len(unassigned_orders)}</b> đơn\n"
    caption += f"  • ⚠️ <b>Đang giao / Giao thất bại:</b> <b>{len(failed_delivery_orders)}</b> đơn\n"
    caption += f"  • ✅ <b>Đã xử lý xong (Thành công):</b> <b>{len(processed_success_orders)}</b> đơn\n\n"
    
    caption += f"🏆 <b>DANH SÁCH AM CẦN ƯU TIÊN XỬ LÝ GẤP:</b>\n"
    
    # Sắp xếp các AM theo số lượng đơn Nhóm 1 (unassigned) giảm dần
    sorted_ams = sorted(am_stats.items(), key=lambda x: x[1]["unassigned"], reverse=True)
    
    for am_name, stats in sorted_ams:
        if stats["unassigned"] == 0:
            continue
            
        days_plus_count = stats.get("unassigned_2days_plus", 0)
        if days_plus_count > 0:
            caption += f"  • AM <b>{am_name}</b>: <b>{stats['unassigned']}</b> đơn chưa gán (Trong đó chưa gán liên tiếp &gt;=2 ngày: <b>{days_plus_count}</b> đơn | Tổng tồn: {stats['total']} đơn)\n"
        else:
            caption += f"  • AM <b>{am_name}</b>: <b>{stats['unassigned']}</b> đơn chưa gán (Tổng tồn: {stats['total']} đơn)\n"
        
    caption += f"\n👉 Anh/chị AM vào link trên, lọc tên mình ở cột AM tại tab <b>Đơn Chưa Xử Lý Aging</b> để lấy danh sách mã đơn cụ thể và push bưu cục gán chuyến đi giao ngay nhé!\n"
    print(caption)
    
    # 7. Gửi lên GTalk
    print("📡 Đang gửi báo cáo sang GTalk...")
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
    check_unprocessed_orders()
