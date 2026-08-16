import os
import io
import sys
import json
import requests
import gspread
import time
import unicodedata
from datetime import datetime, timedelta
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

SUCCESS_KEYWORDS = ['đã giao/ chuyển trả thành công', 'đã giao/chuyển trả thành công', 'n/a', 'thành công']


def get_baseline_path(date_str):
    return os.path.join(BASE_DIR, f'baseline_aging_15_{date_str}.json')


def is_processed(status):
    return status in ['#N/A', 'n/a'] or any(sk in status.lower() for sk in SUCCESS_KEYWORDS)


def build_lm_status(lm_data):
    lm_header = lm_data[1]
    order_col_idx = lm_header.index("Mã đơn hàng")
    status_col_idx = lm_header.index("Trạng thái")
    lm_status = {}
    for row in lm_data[2:]:
        if len(row) > max(order_col_idx, status_col_idx):
            m_don = row[order_col_idx].strip()
            t_thai = row[status_col_idx].strip()
            if m_don:
                lm_status[m_don] = t_thai
    return lm_status


def build_cocau_map(cocau_data):
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
    return cocau_map


def compute_current_aging_15(aging_data, lm_status, cocau_map):
    """Tính danh sách đơn aging > 15 ngày TẠI THỜI ĐIỂM HIỆN TẠI (giống script 1)."""
    aging_header = aging_data[0]
    ag_order_col = aging_header.index("order_code") if "order_code" in aging_header else aging_header.index("mã đơn")
    ag_bc_col = aging_header.index("bc")
    ag_id_bc_col = aging_header.index("id_bc")
    ag_aging_col = aging_header.index("Aging") if "Aging" in aging_header else aging_header.index("aging")
    ag_group_col = aging_header.index("Nhóm BL")
    ag_am_col = aging_header.index("am_name")

    current_by_am = {}   # am -> [codes]
    current_am_of_code = {}  # code -> am (để biết AM của đơn phát sinh mới)

    for row in aging_data[1:]:
        if len(row) > max(ag_order_col, ag_bc_col, ag_id_bc_col, ag_aging_col, ag_group_col, ag_am_col):
            order_code = row[ag_order_col].strip()
            if not order_code:
                continue

            status = lm_status.get(order_code, '#N/A')
            if is_processed(status):
                continue

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
                raw_am = row[ag_am_col].strip()
                if not raw_am or raw_am == '#N/A' or raw_am == '':
                    bc_name = row[ag_bc_col].strip()
                    id_bc = row[ag_id_bc_col].strip()
                    am_name = cocau_map.get(id_bc, cocau_map.get(bc_name, "Không xác định"))
                else:
                    am_name = raw_am
                am_name = unicodedata.normalize('NFC', am_name)

                current_by_am.setdefault(am_name, []).append(order_code)
                current_am_of_code[order_code] = am_name

    return current_by_am, current_am_of_code


def main():
    # Cho phép truyền ngày cần đối chiếu qua tham số dòng lệnh, mặc định là HÔM QUA
    target_date = None
    for arg in sys.argv[1:]:
        if arg not in ('--force', '--reset'):
            target_date = arg

    explicit_date_given = target_date is not None
    if not target_date:
        target_date = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')

    print(f"⏰ Bắt đầu đối chiếu đơn aging > 15 ngày lúc: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    baseline_file = get_baseline_path(target_date)

    # Nếu không có tham số ngày cụ thể và baseline của "hôm qua" bị thiếu (VD: quên chạy hôm đó),
    # tự động lùi lại tìm baseline gần nhất còn tồn tại (tối đa 7 ngày trước).
    fallback_used = False
    if not os.path.exists(baseline_file) and not explicit_date_given:
        search_date = datetime.strptime(target_date, '%Y-%m-%d')
        for _ in range(7):
            search_date -= timedelta(days=1)
            candidate = search_date.strftime('%Y-%m-%d')
            candidate_file = get_baseline_path(candidate)
            if os.path.exists(candidate_file):
                print(f"⚠️ Không có baseline ngày {target_date} (có thể hôm đó quên chạy script tìm đơn). "
                      f"Tự động dùng baseline gần nhất: ngày {candidate}.")
                target_date = candidate
                baseline_file = candidate_file
                fallback_used = True
                break

    if not os.path.exists(baseline_file):
        print(f"❌ Lỗi: Không tìm thấy file baseline cho ngày {target_date} tại {baseline_file}, "
              f"và cũng không tìm thấy baseline nào trong 7 ngày gần nhất. "
              f"Cần kiểm tra lại script tìm đơn aging có đang chạy đều đặn không.")
        sys.exit(1)

    with open(baseline_file, 'r', encoding='utf-8') as f:
        b_data = json.load(f)
    baseline_orders = b_data.get('baseline_orders', {})
    print(f"ℹ️ Đã tải baseline ngày {target_date} ({sum(len(v) for v in baseline_orders.values())} đơn).")

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

    try:
        gc_client = get_gspread_client(SHEET_KEY)
        sh = gc_client.open_by_key(SHEET_KEY)
        ws_lm = sh.worksheet(SHEET_NAME_LM)
        lm_data = ws_lm.get_all_values()
        ws_aging = sh.worksheet(SHEET_NAME_AGING)
        aging_data = ws_aging.get_all_values()
        ws_cocau = sh.worksheet("Cơ cấu")
        cocau_data = ws_cocau.get_all_values()
    except Exception as e:
        print(f"❌ Lỗi kết nối Google Sheets: {e}")
        sys.exit(1)

    if len(lm_data) < 2:
        print("❌ Lỗi: Sheet 'data LM' không đủ dữ liệu.")
        sys.exit(1)

    try:
        lm_status = build_lm_status(lm_data)
    except ValueError as e:
        print(f"❌ Lỗi: Không tìm thấy cột cần thiết trong 'data LM'. Chi tiết: {e}")
        sys.exit(1)

    cocau_map = build_cocau_map(cocau_data)

    try:
        current_by_am, current_am_of_code = compute_current_aging_15(aging_data, lm_status, cocau_map)
    except ValueError as e:
        print(f"❌ Lỗi: Không tìm thấy cột bắt buộc trong tab Đơn giao aging trên 5 ngày. Chi tiết: {e}")
        sys.exit(1)

    # Tập hợp toàn bộ mã đơn có trong baseline (để tính phát sinh mới)
    baseline_codes_all = set()
    for codes in baseline_orders.values():
        baseline_codes_all.update(codes)

    # 2. Đối chiếu từng AM: đã xử lý / còn tồn (dựa trên baseline)
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
            if is_processed(status):
                mat_list.append(code)
            else:
                con_list.append(code)
                is_assigned = status in ["Đang có chuyến đi trong ngày", "Đã có chuyến đi trong ngày"]
                if is_assigned:
                    gan_count += 1
                else:
                    chua_gan_count += 1

        processed_count = len(mat_list)
        con_count = len(con_list)
        pct = round(processed_count / total_baseline * 100) if total_baseline > 0 else 0

        # Phát sinh mới của AM này: có trong current nhưng KHÔNG có trong baseline
        new_codes = [c for c in current_by_am.get(am, []) if c not in baseline_codes_all]

        am_report[am] = {
            "total": total_baseline,
            "con_count": con_count,
            "processed_count": processed_count,
            "gan_count": gan_count,
            "chua_gan_count": chua_gan_count,
            "con_list": con_list,
            "mat_list": mat_list,
            "new_list": new_codes,
            "pct": pct
        }

    # Các AM chỉ xuất hiện ở "phát sinh mới" (không có trong baseline hôm qua)
    for am, codes in current_by_am.items():
        if am not in am_report:
            new_codes = [c for c in codes if c not in baseline_codes_all]
            if new_codes:
                am_report[am] = {
                    "total": 0, "con_count": 0, "processed_count": 0,
                    "gan_count": 0, "chua_gan_count": 0,
                    "con_list": [], "mat_list": [], "new_list": new_codes, "pct": 0
                }

    total_new = sum(len(r["new_list"]) for r in am_report.values())
    total_processed = sum(r["processed_count"] for r in am_report.values())
    total_con = sum(r["con_count"] for r in am_report.values())

    # 3. Format GTalk Caption
    today_str = datetime.now().strftime('%d/%m/%Y')
    current_time = datetime.now().strftime('%H:%M')
    target_date_str = datetime.strptime(target_date, '%Y-%m-%d').strftime('%d/%m/%Y')

    caption = f"📊 <b>ĐỐI CHIẾU ĐƠN AGING &gt; 15 NGÀY (baseline ngày {target_date_str})</b>\n"
    if fallback_used:
        caption += f"⚠️ <i>Lưu ý: không có baseline hôm qua, đang so sánh với baseline gần nhất tìm được ({target_date_str}).</i>\n"
    caption += f"⏱️ <b>Mốc cập nhật:</b> {current_time} ngày {today_str}\n"
    caption += f"✅ Đã xử lý: <b>{total_processed}</b>  |  🕒 Còn tồn: <b>{total_con}</b>  |  🆕 Phát sinh mới: <b>{total_new}</b>\n\n"

    if not am_report:
        caption += "🎉 Không có dữ liệu đối chiếu.\n"
    else:
        sorted_ams = sorted(am_report.items(), key=lambda x: (x[1]['con_count'] + len(x[1]['new_list'])), reverse=True)
        for am, r in sorted_ams:
            caption += f"👤 AM <b>{am}</b>:\n"
            if r["total"] > 0:
                caption += f"  • Đã xử lý: <b>{r['processed_count']}</b>/{r['total']} đơn (Đạt {r['pct']}%)\n"
                caption += f"  • Còn tồn: <b>{r['con_count']}</b> đơn (Đã gán: {r['gan_count']} | Chưa gán: {r['chua_gan_count']})\n"
            if r["new_list"]:
                caption += f"  • 🆕 Phát sinh mới: <b>{len(r['new_list'])}</b> đơn\n"
                caption += "    " + ", ".join([f"<code>{c}</code>" for c in r['new_list']]) + "\n"
            caption += "\n"

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