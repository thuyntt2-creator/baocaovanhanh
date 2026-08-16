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
SHEET_KEY = '1WCzgao34cA_SttyB9ytHfE1qKTNl_3iFqDbEfw3lbyU'  # Main Sheet (Aging >5 ngày)
SHEET_NAME = 'data LM'

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
    print(f"⏰ Bắt đầu đối chiếu đơn aging > 15 ngày lúc: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    # 1. Connect to Google Sheets
    try:
        gc_client = get_gspread_client(SHEET_KEY)
        sh = gc_client.open_by_key(SHEET_KEY)
        ws = sh.worksheet(SHEET_NAME)
        lm_data = ws.get_all_values()
    except Exception as e:
        print(f"❌ Lỗi kết nối Google Sheets: {e}")
        sys.exit(1)
        
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
        
    # Map Mã đơn hàng -> Trạng thái hiện tại
    lm_status = {}
    for row in lm_data[2:]:
        if len(row) > max(order_col_idx, status_col_idx):
            m_don = row[order_col_idx].strip()
            t_thai = row[status_col_idx].strip()
            if m_don:
                lm_status[m_don] = t_thai
                
    # 2. Load or generate baseline
    baseline_file = os.path.join(BASE_DIR, 'baseline_aging_15.json')
    today_ymd = datetime.now().strftime('%Y-%m-%d')
    baseline_orders = None
    
    if os.path.exists(baseline_file):
        try:
            with open(baseline_file, 'r', encoding='utf-8') as f:
                b_data = json.load(f)
                if b_data.get('date') == today_ymd:
                    baseline_orders = b_data.get('baseline_orders')
                    print(f"ℹ️ Đã tải baseline đơn aging > 15 ngày cho ngày {today_ymd} từ file.")
        except Exception as e:
            print(f"⚠️ Lỗi đọc file baseline: {e}")

    if not baseline_orders:
        print(f"⚠️ Chưa có baseline cho ngày {today_ymd} hoặc file bị lỗi. Tiến hành tự động tạo baseline mới từ Google Sheet...")
        try:
            # We already have sh, now read other worksheets
            print("📖 Đọc dữ liệu từ tab 'Đơn giao aging trên 5 ngày'...")
            ws_aging = sh.worksheet('Đơn giao aging trên 5 ngày')
            aging_data = ws_aging.get_all_values()
            
            print("📖 Đọc dữ liệu từ tab 'CoCauVung' / 'Cơ cấu'...")
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
        except Exception as e:
            print(f"❌ Lỗi đọc dữ liệu sheet để tạo baseline: {e}")
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
                    
        # Parse Đơn giao aging trên 5 ngày
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
        
        baseline_orders = {}
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
                    
                    if am_name not in baseline_orders:
                        baseline_orders[am_name] = []
                    baseline_orders[am_name].append(order_code)
                    
        # Save to JSON
        try:
            baseline_data = {
                "date": today_ymd,
                "baseline_orders": baseline_orders
            }
            with open(baseline_file, 'w', encoding='utf-8') as f:
                json.dump(baseline_data, f, ensure_ascii=False, indent=2)
            print(f"✅ Đã tự động tạo và lưu baseline mới cho ngày {today_ymd}!")
        except Exception as e:
            print(f"⚠️ Lỗi lưu file baseline: {e}")

    # 3. Process and compare
    success_keywords = ['đã giao/ chuyển trả thành công', 'đã giao/chuyển trả thành công', 'n/a', 'thành công']
    
    am_report = {}
    for am, codes in baseline_orders.items():
        total_baseline = len(codes)
        if total_baseline == 0:
            continue
        con_list = []
        mat_list = []
        gan_count = 0
        chua_gan_count = 0
        
        for code in codes:
            status = lm_status.get(code, '#N/A')
            is_processed = status in ['#N/A', 'n/a'] or any(sk in status.lower() for sk in success_keywords)
            
            if is_processed:
                mat_list.append(code)
            else:
                con_list.append(code)
                # Check assigned status
                is_assigned = status in ["Đang có chuyến đi trong ngày", "Đã có chuyến đi trong ngày"]
                if is_assigned:
                    gan_count += 1
                else:
                    chua_gan_count += 1
                    
        processed_count = len(mat_list)
        con_count = len(con_list)
        pct = round(processed_count / total_baseline * 100) if total_baseline > 0 else 0
        
        am_report[am] = {
            "total": total_baseline,
            "con_count": con_count,
            "processed_count": processed_count,
            "gan_count": gan_count,
            "chua_gan_count": chua_gan_count,
            "con_list": con_list,
            "mat_list": mat_list,
            "pct": pct
        }

    # 4. Format GTalk Caption
    today_str = datetime.now().strftime('%d/%m/%Y')
    current_time = datetime.now().strftime('%H:%M')
    
    caption = f"📊 <b>THEO DÕI ĐƠN AGING &gt; 15 NGÀY</b>\n"
    caption += f"⏱️ <b>Mốc cập nhật:</b> {current_time} ngày {today_str}\n\n"
    
    if not am_report:
        caption += "🎉 Tuyệt vời! Hiện tại không có đơn hàng tồn đọng Aging > 15 ngày nào cần theo dõi.\n"
    else:
        # Sort AMs by count descending
        sorted_ams = sorted(am_report.items(), key=lambda x: x[1]['total'], reverse=True)
        for am, r in sorted_ams:
            caption += f"👤 AM <b>{am}</b>:\n"
            caption += f"  • Còn tồn: <b>{r['con_count']}</b> đơn (Trong đó: Đã gán: {r['gan_count']} | Chưa gán: {r['chua_gan_count']})\n"
            caption += f"  • Đã xử lý: <b>{r['processed_count']}</b>/{r['total']} đơn (Đạt {r['pct']}%)\n"
            
            # Details of codes
            con_str = ", ".join([f"<code>{c}</code>" for c in r['con_list']]) if r['con_list'] else "Không có đơn nào"
            mat_str = ", ".join([f"<code>{c}</code>" for c in r['mat_list']]) if r['mat_list'] else "Không có đơn nào"
            
            caption += f"  • Chi tiết còn tồn: {con_str}\n"
            caption += f"  • Chi tiết đã xử lý: {mat_str}\n\n"
        
    print("=== NỘI DUNG TIN NHẮN ===")
    print(caption)
    
    # 4. Send to GTalk Group
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
    main()
