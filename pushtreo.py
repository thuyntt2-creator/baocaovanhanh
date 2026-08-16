# -*- coding: utf-8 -*-
"""
Cảnh báo đơn TREO LUÂN CHUYỂN >= 24h (Chưa đóng kiện, tồn đọng >= 24h),
tách riêng theo từng AM và gửi vào ĐÚNG KÊNH GTALK RIÊNG của AM đó.
 
Port lại logic từ Google Apps Script `alertStuckLuanchuyenToAM`
(đọc CơCấuVùng + LuânChuyểnMới + LogNhacNho), nhưng gửi qua GTalk
theo đúng pattern initiate-upload / send-message bạn đang dùng cho NVPTT.

Cách chạy xem trước (Dry-Run, KHÔNG tự gửi tin):
    python pushtreolc.py

Cách chạy gửi thật:
    python pushtreolc.py --force-send
"""
 
import os
import sys
import time
import json
import unicodedata
from datetime import datetime
 
import requests
import gspread
from google.oauth2.service_account import Credentials
 
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

sys.stdout.reconfigure(encoding='utf-8')
 
# ========================= CẤU HÌNH =========================
GOOGLE_SHEET_KEY = "1MjLW8NbD5ZjoOdd90myGv0i1NGAtlvScxebfAXMM1j8"  # Link Google Sheet đơn luân chuyển trễ
 
TAB_COCAU = "Cơ cấu"
TAB_LUANCHUYEN = "stuck"
TAB_LOG = "LogNhacNho"
 
TON_THRESHOLD_HOURS = 24  # Lọc chỉ push đơn tồn đọng >= 24h
 
GTALK_OA_TOKEN = os.environ.get("GTALK_OA_TOKEN") or "2077276776281051136:8hMHvBBU8qXKps3mLPzgKBucPLSQPg3Y"
DEFAULT_CHANNEL_ID = "2076974545807159296"  # fallback nếu AM chưa có kênh riêng trong map
 
AM_CHANNEL_MAP = {
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
 
# Đặt True để CHỈ gửi thử 1 AM đầu tiên (test payload GTalk trước khi bắn hết cho tất cả AM)
TEST_MODE = False

# ============ CHẾ ĐỘ TEST & AN TOÀN ============
# Mặc định ENABLE_SEND = False (Dry-Run: chỉ in thông tin xem trước, KHÔNG tự ý gửi vào GTalk).
# Khi sẵn sàng gửi thật cho toàn bộ AM, đổi ENABLE_SEND = True (hoặc truyền tham số --force-send từ CLI).
ENABLE_SEND = False  
# =============================================================
 
 
def get_credentials_path():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    candidates = [
        os.path.join(script_dir, "credentials.json"),
        r"C:\Users\lap4all\Desktop\Backlog_Automation\credentials.json",
        "credentials.json",
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    return None
 
 
def open_sheet():
    json_path = get_credentials_path()
    if not json_path:
        raise FileNotFoundError("❌ Không tìm thấy credentials.json.")
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_file(json_path, scopes=scopes)
    gc = gspread.authorize(creds)
    return gc.open_by_key(GOOGLE_SHEET_KEY)
 
 
def col_idx(headers, keyword):
    kw = keyword.lower().strip()
    for i, h in enumerate(headers):
        if h and kw in str(h).lower().strip():
            return i
    return -1
 
 
def norm(s):
    return unicodedata.normalize("NFC", str(s).strip()).lower()
 
 
def get_worksheet_flexible(sh, possible_titles):
    all_ws = sh.worksheets()
    title_map = {w.title.lower().replace(" ", "").replace("_", ""): w for w in all_ws}
    
    for candidate in possible_titles:
        key = candidate.lower().replace(" ", "").replace("_", "")
        if key in title_map:
            return title_map[key]
            
    available = [w.title for w in all_ws]
    raise ValueError(f"Không tìm thấy tab phù hợp với các tên: {possible_titles}. Các tab hiện có trong Google Sheet: {available}")
 
 
def load_bc_map(sh):
    """Cơ cấu / CơCấuVùng: warehouse_id -> Tên bưu cục"""
    ws = get_worksheet_flexible(sh, [TAB_COCAU, "Cơ cấu", "CơCấuVùng", "CoCauVung", "cơ cấu", "Cơ cấu vùng_new"])
    values = ws.get_all_values()
    if not values:
        raise ValueError(f"Tab '{ws.title}' rỗng.")
 
    headers = values[0]
    i_id = col_idx(headers, "warehouse_id")
    if i_id == -1: i_id = col_idx(headers, "id_bc")
    if i_id == -1: i_id = col_idx(headers, "mã bưu cục")
    if i_id == -1: i_id = 0
 
    i_name = col_idx(headers, "bưu cục")
    if i_name == -1: i_name = col_idx(headers, "bc")
    if i_name == -1: i_name = 1
 
    bc_map = {}
    for row in values[1:]:
        if len(row) <= max(i_id, i_name):
            continue
        wid = str(row[i_id]).strip()
        if wid:
            bc_map[wid] = str(row[i_name]).strip()
    return bc_map
 
 
def parse_ton_hours(val):
    """'120_gio' hoặc '150' -> float giờ. Lỗi -> 0."""
    s = str(val).strip()
    if not s:
        return 0.0
    try:
        return float(s.split("_")[0])
    except ValueError:
        return 0.0
 
 
def get_lc_aging_group(ton_str, ton_hours):
    s = str(ton_str).strip().lower()
    if '192' in s and '120' not in s:
        return "192h+ (Trên 8 ngày)"
    elif '120_192' in s or (120 <= ton_hours < 192):
        return "120 - 192h (5 - 8 ngày)"
    elif '72_96' in s or '96_120' in s or (72 <= ton_hours < 120):
        return "72 - 120h (3 - 5 ngày)"
    elif '36_48' in s or '48_72' in s or (36 <= ton_hours < 72):
        return "36 - 72h (1.5 - 3 ngày)"
    elif '24_36' in s or (24 <= ton_hours < 36):
        return "24 - 36h (1 - 1.5 ngày)"
    elif ton_hours < 24:
        return "Dưới 24h (< 1 ngày)"
    return "Khác"
 
 
def load_stuck_orders(sh, bc_map):
    """
    Tải tất cả các đơn treo luân chuyển từ tab 'stuck' / 'LuânChuyểnMới'.
    Trả về: { am_name: { bc_name: [ {"code": ma_don, "group": grp, "ton_h": ton_val}, ... ] } }
    """
    ws = get_worksheet_flexible(sh, [TAB_LUANCHUYEN, "stuck", "LuânChuyểnMới", "Luân chuyển mới", "Luân chuyển"])
    values = ws.get_all_values()
    if not values:
        raise ValueError(f"Tab '{ws.title}' rỗng.")
 
    headers = values[0]
    i_madon = col_idx(headers, "mã đơn hàng")
    if i_madon == -1: i_madon = col_idx(headers, "mã vận đơn")
    if i_madon == -1: i_madon = 1
 
    i_status = col_idx(headers, "trạng thái")
    if i_status == -1: i_status = 4
 
    i_ton = col_idx(headers, "thời gian tồn đọng")
    if i_ton == -1: i_ton = 5
 
    i_bc = col_idx(headers, "warehouse_name")
    if i_bc == -1: i_bc = col_idx(headers, "bưu cục")
    if i_bc == -1: i_bc = col_idx(headers, "mã bưu cục")
    if i_bc == -1: i_bc = 6
 
    i_am = col_idx(headers, "am_name")
    if i_am == -1: i_am = col_idx(headers, "am")
    if i_am == -1: i_am = 8
 
    am_backlog = {}          # am -> bc -> [order_dict,...]
    current_stuck_map = {}   # ma_don -> số lần nhắc mới
    total_stuck = 0
 
    for row in values[1:]:
        if len(row) <= max(i_madon, i_status, i_ton, i_bc, i_am):
            continue
 
        ma_don = str(row[i_madon]).strip()
        status = str(row[i_status]).strip()
        ton_raw = str(row[i_ton]).strip()
        ton_val = parse_ton_hours(ton_raw)
        id_bc = str(row[i_bc]).strip()
        am_name = str(row[i_am]).strip()
 
        if not am_name or am_name.lower() in ("n/a", "grand total", "#n/a"):
            am_name = "Chưa gán AM"
 
        # Lọc bỏ đơn đã đóng kiện hoặc mã đơn rỗng
        if status == "Đã đóng kiện" or not ma_don:
            continue
 
        # Nếu có ngưỡng TON_THRESHOLD_HOURS > 0 thì lọc theo giờ (ví dụ >= 24h)
        if TON_THRESHOLD_HOURS > 0 and ton_val < TON_THRESHOLD_HOURS:
            continue
 
        grp = get_lc_aging_group(ton_raw, ton_val)
 
        total_stuck += 1
        current_stuck_map[ma_don] = current_stuck_map.get(ma_don, 0) + 1
 
        bc_name = bc_map.get(id_bc, id_bc or "Không rõ BC")
 
        am_backlog.setdefault(am_name, {}).setdefault(bc_name, []).append({
            "code": ma_don,
            "group": grp,
            "ton_h": ton_val
        })
 
    return am_backlog, current_stuck_map, total_stuck
 
 
def update_log_sheet(sh, current_stuck_map):
    try:
        ws = get_worksheet_flexible(sh, [TAB_LOG, "LogNhacNho", "Log Nhắc Nhở", "Log_Nhac_Nho"])
    except ValueError:
        ws = sh.add_worksheet(title=TAB_LOG, rows=max(len(current_stuck_map) + 10, 20), cols=2)
 
    ws.clear()
    data = [["Mã đơn hàng", "Số lần nhắc"]]
    for ma_don, count in current_stuck_map.items():
        data.append([ma_don, count])
    ws.update(data, value_input_option="USER_ENTERED")
 
 
def get_channel_for_am(am_name):
    target = norm(am_name)
    channel_map = dict(AM_CHANNEL_MAP)
    json_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "am_channel_map.json")
    if os.path.exists(json_path):
        try:
            with open(json_path, "r", encoding="utf-8") as f:
                channel_map.update(json.load(f))
        except Exception:
            pass

    for name, channel_id in channel_map.items():
        if norm(name) == target:
            return channel_id
    print(f"⚠️ Không tìm thấy kênh GTalk riêng cho AM '{am_name}' — dùng kênh mặc định.")
    return DEFAULT_CHANNEL_ID
 
 
def build_messages_for_am(am_name, bc_dict, rank_map, total_ams_region, today_str):
    """
    Tạo tin nhắn phân theo mốc thời gian tồn đọng (192h+, 120-192h, 72-120h, 36-72h, 24-36h, Dưới 24h).
    bc_dict: { bc_name: [ {"code": ma_don, "group": grp, "ton_h": ton_h}, ... ] }
    """
    total_orders = sum(len(orders) for orders in bc_dict.values())
    rank_val = rank_map.get(am_name, rank_map.get(norm(am_name), "-"))
    
    header = f"⚠️ <b>ĐƠN TREO LUÂN CHUYỂN PHÂN THEO TUỔI TỒN (≥ 24H)</b>\n"
    header += f"👤 AM: <b>{am_name}</b> — Total {total_orders} đơn\n"
    header += f"📊 <b>Đứng thứ:</b> {rank_val}/{total_ams_region} toàn vùng về số đơn tồn luân chuyển\n"
    header += f"📅 Ngày: {today_str}\n"
    header += "━━━━━━━━━━━━━━━━━━━\n\n"
 
    GROUP_ORDER = [
        "192h+ (Trên 8 ngày)",
        "120 - 192h (5 - 8 ngày)",
        "72 - 120h (3 - 5 ngày)",
        "36 - 72h (1.5 - 3 ngày)",
        "24 - 36h (1 - 1.5 ngày)",
        "Dưới 24h (< 1 ngày)",
        "Khác"
    ]
 
    by_aging = {}
    for bc_name, order_list in bc_dict.items():
        for o in order_list:
            grp = o["group"]
            by_aging.setdefault(grp, {}).setdefault(bc_name, []).append(o["code"])
 
    messages = []
    current_msg = header
    MAX_LEN = 3500
 
    for grp in GROUP_ORDER:
        if grp not in by_aging:
            continue
        
        grp_bcs = by_aging[grp]
        grp_total = sum(len(codes) for codes in grp_bcs.values())
        
        if grp.startswith("192h"):
            grp_header = f"🔴 <b>{grp}</b> ({grp_total} đơn):\n"
        elif grp.startswith("120"):
            grp_header = f"🟠 <b>{grp}</b> ({grp_total} đơn):\n"
        elif grp.startswith("72"):
            grp_header = f"🟡 <b>{grp}</b> ({grp_total} đơn):\n"
        elif grp.startswith("36"):
            grp_header = f"🔵 <b>{grp}</b> ({grp_total} đơn):\n"
        elif grp.startswith("24"):
            grp_header = f"🟣 <b>{grp}</b> ({grp_total} đơn):\n"
        else:
            grp_header = f"⚪ <b>{grp}</b> ({grp_total} đơn):\n"
 
        bc_blocks = []
        for bc_name, code_list in grp_bcs.items():
            code_tags = [f"<code>{c}</code>" for c in code_list]
            bc_blocks.append(f"  • <b>{bc_name}</b> ({len(code_list)} đơn):\n    ➔ " + ", ".join(code_tags))
            
        full_grp_str = grp_header + "\n".join(bc_blocks) + "\n\n"
 
        if len(current_msg) + len(full_grp_str) <= MAX_LEN:
            current_msg += full_grp_str
        else:
            if len(current_msg) > len(header):
                messages.append(current_msg.strip())
                current_msg = f"<b>ĐƠN TREO LUÂN CHUYỂN (Tiếp theo) - AM {am_name}</b>\n━━━━━━━━━━━━━━━━━━━\n\n"
            
            current_msg += grp_header
            for block in bc_blocks:
                if len(current_msg) + len(block) + 2 > MAX_LEN:
                    messages.append(current_msg.strip())
                    current_msg = f"<b>ĐƠN TREO LUÂN CHUYỂN (Tiếp theo) - AM {am_name} ({grp})</b>\n━━━━━━━━━━━━━━━━━━━\n\n" + block + "\n\n"
                else:
                    current_msg += block + "\n\n"
 
    if current_msg.strip():
        messages.append(current_msg.strip())
 
    return messages
 
 
def send_gtalk_text_message(channel_id, text, oa_token=GTALK_OA_TOKEN):
    """
    Gửi tin nhắn TEXT thuần (không ảnh) qua GTalk.
    """
    payload = {
        "channelId": str(channel_id),
        "clientMsgId": str(int(datetime.now().timestamp() * 1000)),
        "content": {
            "parseMode": "HTML",
            "text": text
        },
        "oaToken": oa_token
    }
    try:
        r = requests.post("https://mbff.ghn.vn/api/gtalk/send-message", json=payload, timeout=15, verify=False)
        try:
            res_json = r.json()
        except Exception:
            res_json = {}

        if r.status_code == 200 and res_json.get("errorCode") == "success":
            return True

        # Nếu kênh riêng bị lỗi CHANNEL_NOT_BELONG_TO_OA, fallback gửi qua kênh mặc định
        if isinstance(res_json, dict) and res_json.get("error", {}).get("errorCode") == "CHANNEL_NOT_BELONG_TO_OA" and str(channel_id) != str(DEFAULT_CHANNEL_ID):
            print(f"⚠️ Kênh {channel_id} chưa gán OA — tự động gửi về kênh mặc định ({DEFAULT_CHANNEL_ID})...")
            return send_gtalk_text_message(DEFAULT_CHANNEL_ID, text, oa_token)

        print(f"❌ Gửi GTalk lỗi (channel {channel_id}): HTTP {r.status_code} — {r.text}")
        return False
    except Exception as e:
        print(f"❌ Exception khi gửi GTalk (channel {channel_id}): {e}")
        return False
 
 
def main():
    today_str = datetime.now().strftime("%d/%m/%Y")
 
    print("🔌 Đang kết nối Google Sheet...")
    sh = open_sheet()
 
    print(f"📥 Đọc '{TAB_COCAU}'...")
    bc_map = load_bc_map(sh)
 
    print(f"📥 Đọc tab luân chuyển ({GOOGLE_SHEET_KEY})...")
    am_backlog, current_stuck_map, total_stuck = load_stuck_orders(sh, bc_map)
 
    print(f"💾 Cập nhật log nhắc nhở ('{TAB_LOG}')...")
    update_log_sheet(sh, current_stuck_map)
 
    if total_stuck == 0:
        print("✅ Không có đơn luân chuyển bị treo nào (>= 24h) trong file. Không gửi tin.")
        return
 
    print(f"📊 Tổng {total_stuck} đơn treo (>= 24h), thuộc {len(am_backlog)} AM. Bắt đầu gửi từng kênh GTalk...")
 
    channel_map = dict(AM_CHANNEL_MAP)
    json_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "am_channel_map.json")
    if os.path.exists(json_path):
        try:
            with open(json_path, "r", encoding="utf-8") as f:
                channel_map.update(json.load(f))
        except Exception:
            pass

    all_region_ams = list(channel_map.keys())
    ranked_ams = sorted(all_region_ams, key=lambda am: sum(len(orders) for orders in am_backlog.get(am, {}).values()), reverse=True)
    rank_map = {norm(am): i + 1 for i, am in enumerate(ranked_ams)}
    total_ams_region = len(all_region_ams)

    items = list(am_backlog.items())
    if TEST_MODE:
        items = items[:1]
        print("🧪 TEST_MODE=True -> chỉ gửi thử AM đầu tiên.")
 
    is_real_send = ENABLE_SEND or '--force-send' in sys.argv

    for i, (am_name, bc_dict) in enumerate(items, start=1):
        channel_id = get_channel_for_am(am_name)
        messages = build_messages_for_am(am_name, bc_dict, rank_map, total_ams_region, today_str)

        total_nv = sum(len(v) for v in bc_dict.values())
        
        if not is_real_send:
            print(f"👀 [DRY-RUN - KHÔNG GỬI TIN THẬT] [{i}/{len(items)}] AM '{am_name}' ({total_nv} đơn - {len(messages)} tin nhắn)")
            continue

        print(f"[{i}/{len(items)}] 📤 Đang gửi AM '{am_name}' ({total_nv} đơn - {len(messages)} tin nhắn)...", end="")
        am_ok = True
        for msg in messages:
            if not send_gtalk_text_message(channel_id, msg):
                am_ok = False
                break
            time.sleep(0.3)

        if am_ok:
            print(" [THÀNH CÔNG]")
        else:
            print(" [THẤT BẠI]")

        if i < len(items):
            time.sleep(1)

    print("\n🎉 Đã hoàn tất xử lý.")
 
 
if __name__ == "__main__":
    main()
