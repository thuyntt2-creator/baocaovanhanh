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
from dotenv import load_dotenv

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

# Google Sheets Config
SHEET_KEY = '1WCzgao34cA_SttyB9ytHfE1qKTNl_3iFqDbEfw3lbyU'  # Main Sheet
SHEET_NAME_LM = 'data LM'
SHEET_NAME_AGING = 'Đơn giao aging trên 5 ngày'

# Load environment configuration if available
env_path = r"c:\Users\lap4all\Desktop\New folder\.env"
if os.path.exists(env_path):
    load_dotenv(dotenv_path=env_path, override=True)
else:
    load_dotenv()

# GTalk Config
GTALK_OA_TOKEN = os.environ.get("AGING_15_GTALK_OA_TOKEN") or "2067164759497973760:RfgqBJY4QtV18udu1wMEfhmoRNI4hgBv"
GTALK_CHANNEL_ID = os.environ.get("AGING_15_GTALK_CHANNEL_ID") or "2067164759710552066"

scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']


def get_gspread_client(spreadsheet_id=SHEET_KEY):
    cred_paths = [
        JSON_FILE,
        os.path.join(BASE_DIR, 'credentials.json'),
        r'C:\Users\lap4all\Documents\Auto report\credentials.json',
        r'C:\Users\lap4all\Desktop\Backlog_Automation\credentials.json',
    ]
    for cred_path in cred_paths:
        if os.path.isfile(cred_path):
            try:
                creds = Credentials.from_service_account_file(cred_path, scopes=scopes)
                gc = gspread.authorize(creds)
                if spreadsheet_id:
                    gc.open_by_key(spreadsheet_id)
                return gc
            except Exception as e:
                print(f"⚠️ Service account ({cred_path}) không có quyền: {e}. Đang chuyển sang authorized_user.json...", flush=True)

    auth_user_paths = [
        os.path.join(BASE_DIR, 'authorized_user.json'),
        r'C:\Users\lap4all\Documents\Auto report\authorized_user.json',
        r'C:\Users\lap4all\Desktop\Backlog_Automation\authorized_user.json',
        'authorized_user.json'
    ]
    for auth_user_file in auth_user_paths:
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

    oauth_file = os.path.join(BASE_DIR, 'credentials_oauth.json')
    if os.path.exists(oauth_file):
        try:
            gc = gspread.oauth(
                credentials_filename=oauth_file,
                authorized_user_filename=os.path.join(BASE_DIR, 'authorized_user.json')
            )
            if spreadsheet_id:
                gc.open_by_key(spreadsheet_id)
            return gc
        except Exception as e:
            pass

    raise PermissionError("Không thể xác thực Google Sheets bằng credentials.json hoặc authorized_user.json")


def main():
    print(f"⏰ Bắt đầu tổng hợp chi tiết đơn aging > 15 ngày lúc: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    # 1. Connect to Google Sheets
    try:
        gc_client = get_gspread_client(SHEET_KEY)
        sh = gc_client.open_by_key(SHEET_KEY)
        
        # Read data LM
        ws_lm = sh.worksheet(SHEET_NAME_LM)
        lm_data = ws_lm.get_all_values()
        
        # Read Đơn giao aging trên 5 ngày
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
    except ValueError as e:
        print(f"❌ Lỗi: Không tìm thấy cột bắt buộc trong tab Đơn giao aging trên 5 ngày. Chi tiết: {e}")
        sys.exit(1)
        
    success_keywords = ['đã giao/ chuyển trả thành công', 'đã giao/chuyển trả thành công', 'n/a', 'thành công']
    aging_15_orders_by_am = {}
    total_aging_15_count = 0
    
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
                
            # Check if aging > 15 days
            group_val = row[ag_group_col].strip()
            aging_val = row[ag_aging_col].strip()
            
            is_over_15 = False
            try:
                if aging_val and float(aging_val) > 15.0:
                    is_over_15 = True
            except ValueError:
                pass
            if not is_over_15 and '(k)' in group_val.lower():
                is_over_15 = True
                
            if is_over_15:
                # Match found! Map AM name
                raw_am = row[ag_am_col].strip()
                if not raw_am or raw_am == '#N/A' or raw_am == '':
                    bc_name = row[ag_bc_col].strip()
                    id_bc = row[ag_id_bc_col].strip()
                    am_name = cocau_map.get(id_bc, cocau_map.get(bc_name, "Không xác định"))
                else:
                    am_name = raw_am
                am_name = unicodedata.normalize('NFC', am_name)
                
                if am_name not in aging_15_orders_by_am:
                    aging_15_orders_by_am[am_name] = []
                    
                aging_15_orders_by_am[am_name].append({
                    "code": order_code,
                    "bc": row[ag_bc_col].strip(),
                    "status": status,
                    "aging_days": aging_val
                })
                total_aging_15_count += 1

    # 3.5. Update or save baseline for follow_aging_15
    baseline_file = os.path.join(BASE_DIR, 'baseline_aging_15.json')
    today_ymd = datetime.now().strftime('%Y-%m-%d')
    force_reset = '--force' in sys.argv or '--reset' in sys.argv
    
    baseline_orders = {}
    for am, orders in aging_15_orders_by_am.items():
        baseline_orders[am] = [o["code"] for o in orders]
        
    write_baseline = True
    if os.path.exists(baseline_file) and not force_reset:
        try:
            with open(baseline_file, 'r', encoding='utf-8') as f:
                existing_data = json.load(f)
                if existing_data.get('date') == today_ymd:
                    write_baseline = False
                    print(f"ℹ️ Baseline cho ngày hôm nay ({today_ymd}) đã tồn tại. Bỏ qua ghi đè để tránh làm giảm baseline trong ngày. Sử dụng --force để reset.")
        except Exception as e:
            print(f"⚠️ Lỗi đọc file baseline cũ: {e}. Tiến hành ghi đè...")
            
    if write_baseline:
        baseline_data = {
            "date": today_ymd,
            "baseline_orders": baseline_orders
        }
        try:
            with open(baseline_file, 'w', encoding='utf-8') as f:
                json.dump(baseline_data, f, ensure_ascii=False, indent=2)
            print(f"✅ Đã lưu baseline đơn aging > 15 ngày cho ngày {today_ymd} thành công!")
        except Exception as e:
            print(f"❌ Lỗi ghi file baseline: {e}")

    # 4. Format GTalk Notification Message
    today_str = datetime.now().strftime('%d/%m/%Y')
    current_time = datetime.now().strftime('%H:%M')
    
    caption = f"<b>CHI TIẾT ĐƠN AGING &gt; 15 NGÀY</b>\n"
    caption += f"<b>Mốc cập nhật:</b> {current_time} ngày {today_str}\n"
    caption += f"<b>Tổng số đơn tồn &gt; 15 ngày:</b> <b>{total_aging_15_count}</b> đơn\n"
    caption += f"🔗 <b>Xem chi tiết báo cáo:</b> <a href=\"https://docs.google.com/spreadsheets/d/{SHEET_KEY}/edit#gid=1560830058\"><b>Click xem</b></a>\n\n"
    
    if total_aging_15_count > 0:
        caption += f"🏆 <b>DANH SÁCH CHI TIẾT THEO AM:</b>\n"
        # Sort AMs by count descending
        sorted_ams = sorted(aging_15_orders_by_am.items(), key=lambda x: len(x[1]), reverse=True)
        
        for am, orders in sorted_ams:
            caption += f"AM <b>{am}</b>: <b>{len(orders)}</b> đơn\n"
            for o in orders:
                caption += f"<code>{o['code']}</code>\n"
    else:
        caption += f"🎉 Tuyệt vời! Hiện tại không có đơn hàng tồn đọng Aging > 15 ngày nào."

    print("=== NỘI DUNG TIN NHẮN ===")
    print(caption)
    
    # 5. Send to GTalk Group
    print("📡 Đang gửi báo cáo chi tiết sang GTalk...")
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
