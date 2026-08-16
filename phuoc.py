import os
import io
import sys
import json
import requests
import gspread
import time
import unicodedata
import urllib3
from datetime import datetime
from google.oauth2.service_account import Credentials
from google.oauth2.credentials import Credentials as UserCredentials
from dotenv import load_dotenv

# Tắt cảnh báo SSL rác khi verify=False
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ============ Fix encoding cho Windows/Task Scheduler/PowerShell ============
os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except AttributeError:
    pass

# ============ Paths & Constants ============
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
AM_CHANNEL_MAP_FILE = os.path.join(BASE_DIR, 'am_channel_map.json')

SHEET_KEY = '1WCzgao34cA_SttyB9ytHfE1qKTNl_3iFqDbEfw3lbyU'
SHEET_NAME_LM = 'data LM'
SHEET_NAME_AGING = 'Đơn giao aging trên 5 ngày'

ALL_GROUPS = ['5 - 8 ngày', '>8 - 10 ngày', '>10 - 15 ngày', 'Trên 15 ngày']
ASSIGNED_STATUSES = ["Đang có chuyến đi trong ngày", "Đã có chuyến đi trong ngày"]
MAX_CODES_PER_GROUP = 40

# ============ CHẾ ĐỘ TEST & AN TOÀN ============
# Để an toàn, mặc định ENABLE_SEND = False (Dry-Run: chỉ in thông tin xem trước, KHÔNG tự ý gửi vào GTalk).
# Khi sẵn sàng gửi thật cho toàn bộ AM, đổi ENABLE_SEND = True (hoặc truyền tham số --force-send từ CLI).
ENABLE_SEND = False  

TEST_MODE = False
TEST_AM_NAME = "Trần Văn Phước"
TEST_CHANNEL_ID = "2077277797832024064"

# Nạp file .env linh hoạt
env_path = os.path.join(BASE_DIR, '.env')
if os.path.exists(env_path):
    load_dotenv(dotenv_path=env_path, override=True)
else:
    load_dotenv()

GTALK_OA_TOKEN = os.environ.get("GTALK_OA_TOKEN") or "2077276776281051136:8hMHvBBU8qXKps3mLPzgKBucPLSQPg3Y"

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

SERVICE_ACCOUNT_CANDIDATES = [
    os.path.join(BASE_DIR, 'credentials.json'),
    r'C:\Users\lap4all\Documents\Auto report\credentials.json',
    r'C:\Users\lap4all\Desktop\Backlog_Automation\credentials.json',
    'credentials.json'
]

AUTH_USER_CANDIDATES = [
    os.path.join(BASE_DIR, 'authorized_user.json'),
    r'C:\Users\lap4all\Documents\Auto report\authorized_user.json',
    r'C:\Users\lap4all\Desktop\Backlog_Automation\authorized_user.json',
    'authorized_user.json'
]


def get_gspread_client(spreadsheet_id=None):
    """Thử Service Account trước, fallback sang authorized_user.json nếu bị từ chối quyền."""
    for cred_path in SERVICE_ACCOUNT_CANDIDATES:
        if os.path.isfile(cred_path):
            try:
                creds = Credentials.from_service_account_file(cred_path, scopes=SCOPES)
                gc = gspread.authorize(creds)
                if spreadsheet_id:
                    gc.open_by_key(spreadsheet_id)
                return gc
            except Exception as e:
                print(f"⚠️ Service account ({cred_path}) không có quyền: {e}. Đang chuyển sang authorized_user.json...", flush=True)

    for auth_file in AUTH_USER_CANDIDATES:
        if os.path.exists(auth_file):
            try:
                creds = UserCredentials.from_authorized_user_file(auth_file, scopes=SCOPES)
                gc = gspread.authorize(creds)
                if spreadsheet_id:
                    gc.open_by_key(spreadsheet_id)
                return gc
            except Exception:
                pass

    raise PermissionError("Không thể xác thực Google Sheets bằng credentials.json hoặc authorized_user.json")


def norm_str(s: str) -> str:
    if not s:
        return ""
    return unicodedata.normalize('NFC', str(s).strip())

def get_group(col_k_val: str):
    val = norm_str(col_k_val).lower()
    
    # 1. Kiểm tra ký hiệu phân loại (a, b, c...)
    if any(k in val for k in ['(a)', '(b)', '(c)']):
        return '5 - 8 ngày'
    elif any(k in val for k in ['(d)', '(e)']):
        return '>8 - 10 ngày'
    elif any(k in val for k in ['(f)', '(g)', '(h)', '(i)', '(j)']):
        return '>10 - 15 ngày'
    elif '(k)' in val:
        return 'Trên 15 ngày'
    
    # 2. Khắc phục lỗi: So sánh trực tiếp tên nhóm nếu không có (a), (b)...
    if '5 - 8' in val or '5-8' in val:
        return '5 - 8 ngày'
    elif '>8 - 10' in val or '8-10' in val:
        return '>8 - 10 ngày'
    elif '>10 - 15' in val or '10-15' in val:
        return '>10 - 15 ngày'
    elif 'trên 15' in val or '>15' in val:
        return 'Trên 15 ngày'
        
    return None

DEFAULT_AM_CHANNEL_MAP = {
    "Nguyễn Ngọc Khánh": "2077277510775197696",
    "Nguyễn Duy Long": "2077277718988836864",
    "Lê Thanh Nhựt": "2077277754418147328",
    "Trần Văn Phước": "2077277797832024064",
    "Trần Thị Nhung": "2077277827057745920",
    "Huỳnh Thị Kim Chi": "2077277857186131968",
    "Phan Đình Duy": "2077278383281876992",
    "Phạm Bá Thành Công": "2077277907735883776",
    "Thái Thị Thanh Thư": "2077277934947827712",
    "Nguyễn Thanh Long": "2077277974325506048",
    "Nguyễn Hoàng Phi": "2077278021170323456",
    "Trầm Hữu Tiến": "2077278046487142400",
    "Nguyễn Lê Nguyên Vũ": "2077278095459835904",
    "Lê Văn Trường": "2077278127814696960",
    "Hồng Bích Nga": "2077278157729837056",
    "Lê Minh Đại": "2077278182818799616",
    "Phan Thị Ngọc Diễm": "2079827073949868032",
    "Lê Hồng Minh Tâm": "2079827054540226560",
}

def load_am_channel_map() -> dict:
    merged = {norm_str(k): str(v).strip() for k, v in DEFAULT_AM_CHANNEL_MAP.items()}
    if os.path.exists(AM_CHANNEL_MAP_FILE):
        try:
            with open(AM_CHANNEL_MAP_FILE, 'r', encoding='utf-8') as f:
                raw = json.load(f)
            from_file = {norm_str(k): str(v).strip() for k, v in raw.items()}
            merged.update(from_file)
        except Exception as e:
            print(f"⚠️ Lỗi đọc file am_channel_map.json: {e}")
    return merged

def send_gtalk(channel_id: str, text: str, oa_token: str = None) -> bool:
    token = oa_token or GTALK_OA_TOKEN
    url = "https://mbff.ghn.vn/api/gtalk/send-message"
    client_msg_id = str(int(time.time() * 1000))
    
    payload = {
        "channelId": str(channel_id),
        "clientMsgId": client_msg_id,
        "content": {
            "parseMode": "HTML", 
            "text": text
        },
        "oaToken": token
    }
    headers = {
        "Content-Type": "application/json",
        "User-Agent": "Mozilla/5.0"
    }
    
    try:
        res = requests.post(url, json=payload, headers=headers, timeout=20, verify=False)
        if res.status_code == 200:
            res_data = res.json()
            if res_data.get("errorCode") == "success":
                return True
            print(f"  ❌ GTalk Error ({channel_id}): {res_data}")
            return False
        print(f"  ❌ HTTP Error {res.status_code} ({channel_id}): {res.text}")
        return False
    except Exception as e:
        print(f"  ❌ Connection Error ({channel_id}): {e}")
        return False

def fmt_rate(assigned: int, active: int) -> str:
    if active == 0:
        return "— (0/0)"
    rate = (assigned / active) * 100
    return f"{rate:.1f}% ({assigned}/{active})"

def get_col_index(headers: list, possible_names: list) -> int:
    norm_headers = [norm_str(h).lower() for h in headers]
    for name in possible_names:
        name_clean = norm_str(name).lower()
        if name_clean in norm_headers:
            return norm_headers.index(name_clean)
    raise ValueError(f"Không tìm thấy cột: {possible_names} trong tiêu đề {headers}")

def safe_get_cell(row: list, idx: int) -> str:
    """Lấy giá trị từ dòng an toàn, không sợ IndexError do ô cuối dòng bị rỗng"""
    if 0 <= idx < len(row):
        return norm_str(row[idx])
    return ""

def build_am_messages(am_name, orders_sorted, am_stats, rank_map, total_ams_region, current_time, today_str):
    active_10plus = am_stats['>10 - 15 ngày']['active'] + am_stats['Trên 15 ngày']['active']
    assigned_10plus = am_stats['>10 - 15 ngày']['assigned'] + am_stats['Trên 15 ngày']['assigned']

    header = f"<b>ĐƠN AGING CẦN GÁN</b>\n"
    header += f"<b>AM:</b> {am_name}\n"
    header += f"<b>Mốc cập nhật:</b> {current_time} ngày {today_str}\n"
    header += f"<b>Đứng thứ:</b> {rank_map.get(am_name, '-')}/{total_ams_region} toàn vùng về số đơn tồn\n"
    header += f"<b>Tỷ lệ gán ≥10 ngày:</b> {fmt_rate(assigned_10plus, active_10plus)}\n"
    header += f"  • &gt;10-15 ngày: {fmt_rate(am_stats['>10 - 15 ngày']['assigned'], am_stats['>10 - 15 ngày']['active'])}\n"
    header += f"  • Trên 15 ngày: {fmt_rate(am_stats['Trên 15 ngày']['assigned'], am_stats['Trên 15 ngày']['active'])}\n"
    header += f"<b>Số đơn cần gán ngay:</b> {len(orders_sorted)} đơn\n\n"

    GROUP_ORDER = ['Trên 15 ngày', '>10 - 15 ngày', '>8 - 10 ngày', '5 - 8 ngày']
    by_group = {}
    for o in orders_sorted:
        by_group.setdefault(o["group"], []).append(o)

    messages = []
    current_msg = header
    MAX_LEN = 3500

    for g in GROUP_ORDER:
        if g not in by_group:
            continue
        g_orders = by_group[g]
        g_title = f"<b>{g}</b> ({len(g_orders)} đơn):\n"
        
        # Gom đơn theo Bưu cục (bc)
        by_bc = {}
        for o in g_orders:
            bc_name = o.get('bc') or "Không xác định"
            by_bc.setdefault(bc_name, []).append(o['code'])
            
        bc_blocks = []
        for bc_name, codes in by_bc.items():
            code_tags = [f"<code>{c}</code>" for c in codes]
            bc_blocks.append(f"• <b>{bc_name}</b> ({len(codes)} đơn): " + ", ".join(code_tags))
            
        full_group_str = g_title + "\n".join(bc_blocks) + "\n\n"
        
        if len(current_msg) + len(full_group_str) <= MAX_LEN:
            current_msg += full_group_str
        else:
            if len(current_msg) > len(header):
                messages.append(current_msg.strip())
                current_msg = f"<b>ĐƠN AGING CẦN GÁN (Tiếp theo) - AM {am_name}</b>\n\n"
            
            current_msg += g_title
            chunk_lines = []
            for line in bc_blocks:
                test_str = "\n".join(chunk_lines + [line])
                if len(current_msg) + len(test_str) > MAX_LEN:
                    if chunk_lines:
                        current_msg += "\n".join(chunk_lines) + "\n\n"
                        messages.append(current_msg.strip())
                        current_msg = f"<b>ĐƠN AGING CẦN GÁN (Tiếp theo) - AM {am_name} ({g})</b>\n\n"
                        chunk_lines = [line]
                    else:
                        chunk_lines = [line]
                else:
                    chunk_lines.append(line)
            
            if chunk_lines:
                current_msg += "\n".join(chunk_lines) + "\n\n"

    if current_msg.strip():
        messages.append(current_msg.strip())

    return messages

def main():
    print("🔖 SCRIPT VERSION: v3.5-display-all-codes")
    print(f"⏰ Bắt đầu chạy lúc: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    test_am_filter = None
    test_channel_id = None
    test_oa_token = None

    if TEST_MODE:
        test_am_filter = norm_str(TEST_AM_NAME)
        test_channel_id = TEST_CHANNEL_ID
        print(f"🧪 Đang bật TEST_MODE cho AM '{test_am_filter}'")

    cli_args = [a.strip() for a in sys.argv[1:] if not a.startswith('--') and a.strip()]
    if cli_args:
        test_am_filter = norm_str(cli_args[0])
        print(f"🧪 Nhận tham số CLI AM: '{test_am_filter}'")
    if len(cli_args) > 1:
        test_channel_id = cli_args[1]
    if len(cli_args) > 2:
        test_oa_token = cli_args[2]

    try:
        gc_client = get_gspread_client(spreadsheet_id=SHEET_KEY)
        sh = gc_client.open_by_key(SHEET_KEY)
        lm_data = sh.worksheet(SHEET_NAME_LM).get_all_values()
        aging_data = sh.worksheet(SHEET_NAME_AGING).get_all_values()
        cocau_data = sh.worksheet("Cơ cấu").get_all_values()
    except Exception as e:
        print(f"❌ Lỗi kết nối Google Sheets: {e}")
        sys.exit(1)

    # 1. Tạo Map Cơ Cấu (Chuẩn hóa NFC cho cả ID, BC và AM)
    cocau_map = {}
    for r in cocau_data[1:]:
        if len(r) >= 4:
            id_bc = norm_str(r[0])
            bc_name = norm_str(r[1])
            am_name = norm_str(r[3])
            if id_bc: cocau_map[id_bc] = am_name
            if bc_name: cocau_map[bc_name] = am_name

    # 2. Đọc dữ liệu Tab LM (Kiểm tra xem dòng 0 hay dòng 1 là tiêu đề)
    lm_header_idx = 0
    if len(lm_data) > 1:
        h0 = [norm_str(x).lower() for x in lm_data[0]]
        if not any(k in h0 for k in ["mã đơn hàng", "order_code"]):
            lm_header_idx = 1
            
    lm_header = lm_data[lm_header_idx]
    lm_order_col = get_col_index(lm_header, ["mã đơn hàng", "order_code"])
    lm_status_col = get_col_index(lm_header, ["trạng thái", "status"])

    lm_status = {}
    for row in lm_data[lm_header_idx + 1:]:
        m_don = safe_get_cell(row, lm_order_col)
        st = safe_get_cell(row, lm_status_col)
        if m_don:
            lm_status[m_don] = st

    # 3. Đọc dữ liệu Tab Aging
    aging_header = aging_data[0]
    ag_order_col = get_col_index(aging_header, ["order_code", "mã đơn"])
    ag_bc_col = get_col_index(aging_header, ["bc"])
    ag_id_bc_col = get_col_index(aging_header, ["id_bc"])
    ag_aging_col = get_col_index(aging_header, ["aging"])
    ag_group_col = get_col_index(aging_header, ["nhóm bl"])
    ag_am_col = get_col_index(aging_header, ["am_name"])

    success_keywords = ['đã giao/ chuyển trả thành công', 'đã giao/chuyển trả thành công', 'n/a', 'thành công']
    
    stats = {}
    unassigned_by_am = {}

    for row in aging_data[1:]:
        order_code = safe_get_cell(row, ag_order_col)
        if not order_code:
            continue

        status = lm_status.get(order_code, '#N/A')
        if status in ['#N/A', 'n/a'] or any(sk in status.lower() for sk in success_keywords):
            continue

        group = get_group(safe_get_cell(row, ag_group_col))
        if not group:
            continue

        raw_am = safe_get_cell(row, ag_am_col)
        if not raw_am or raw_am == '#N/A':
            id_bc_val = safe_get_cell(row, ag_id_bc_col)
            bc_val = safe_get_cell(row, ag_bc_col)
            am_name = cocau_map.get(id_bc_val, cocau_map.get(bc_val, "Không xác định"))
        else:
            am_name = raw_am
            
        am_name = norm_str(am_name)

        if am_name not in stats:
            stats[am_name] = {g: {"active": 0, "assigned": 0} for g in ALL_GROUPS}

        is_assigned = status in ASSIGNED_STATUSES
        stats[am_name][group]["active"] += 1
        if is_assigned:
            stats[am_name][group]["assigned"] += 1
        else:
            try:
                aging_float = float(safe_get_cell(row, ag_aging_col))
            except ValueError:
                aging_float = 0.0

            unassigned_by_am.setdefault(am_name, []).append({
                "code": order_code,
                "bc": safe_get_cell(row, ag_bc_col),
                "status": status,
                "aging_days": aging_float,
                "group": group
            })

    push_by_am = unassigned_by_am
    total_push_count = sum(len(v) for v in push_by_am.values())
    print(f"📦 Tổng số đơn chưa gán cần push: {total_push_count} đơn (thuộc {len(push_by_am)} AM)")

    # 4. Xếp hạng & Gửi tin nhắn
    all_region_ams = list(stats.keys())
    ranked_ams = sorted(all_region_ams, key=lambda am: len(push_by_am.get(am, [])), reverse=True)
    rank_map = {am: i + 1 for i, am in enumerate(ranked_ams)}
    total_ams_region = len(all_region_ams)

    am_channel_map = load_am_channel_map()
    today_str = datetime.now().strftime('%d/%m/%Y')
    current_time = datetime.now().strftime('%H:%M')

    sent_count, fail_count = 0, 0

    for am_name, orders in push_by_am.items():
        if test_am_filter and am_name != test_am_filter:
            continue

        am_stats = stats.get(am_name, {g: {"active": 0, "assigned": 0} for g in ALL_GROUPS})
        orders_sorted = sorted(orders, key=lambda o: o["aging_days"], reverse=True)

        messages = build_am_messages(
            am_name=am_name,
            orders_sorted=orders_sorted,
            am_stats=am_stats,
            rank_map=rank_map,
            total_ams_region=total_ams_region,
            current_time=current_time,
            today_str=today_str
        )

        channel_id = test_channel_id or am_channel_map.get(am_name)
        if not channel_id:
            print(f"⚠️ Thiếu channelId cho AM '{am_name}' — Bỏ qua.")
            continue

        # CHÚ Ý: Muốn gửi thật phải bật ENABLE_SEND = True hoặc truyền tham số --force-send
        is_real_send = ENABLE_SEND or '--force-send' in sys.argv

        if not is_real_send:
            print(f"👀 [DRY-RUN - KHÔNG GỬI TIN THẬT] AM {am_name} ({len(orders_sorted)} đơn - {len(messages)} tin nhắn)")
            sent_count += 1
            continue

        print(f"📡 Đang gửi cho AM {am_name} ({len(orders_sorted)} đơn - {len(messages)} tin nhắn)...", end="")
        am_success = True
        for msg in messages:
            if not send_gtalk(channel_id, msg, oa_token=test_oa_token):
                am_success = False
                break
            time.sleep(0.3)

        if am_success:
            print(" [THÀNH CÔNG]")
            sent_count += 1
        else:
            print(" [THẤT BẠI]")
            fail_count += 1

    print("\n=== TỔNG KẾT ===")
    print(f"✅ Gửi thành công: {sent_count} AM")
    print(f"❌ Gửi thất bại  : {fail_count} AM")

if __name__ == "__main__":
    main()
