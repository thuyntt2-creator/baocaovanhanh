# -*- coding: utf-8 -*-
"""
Script: push_am_tts_pickup_individual.py
Reads TTS pickup orders from the 'raw' sheet of the OPR spreadsheet,
maps unassigned orders to AMs, ranks AMs by pending order count,
formats individual alert messages containing region summary, AM rank,
and detailed post offices with tracking numbers (Mã vận đơn - MVĐ),
and sends private messages to each AM via GTalk API.
"""

import os
import sys
import time
import json
import argparse
import requests
import gspread
import pandas as pd
import unicodedata
import html
from google.oauth2.service_account import Credentials
from google.oauth2.credentials import Credentials as UserCredentials
from datetime import datetime
from dotenv import load_dotenv

import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Configure output encoding for Windows Console
os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8', write_through=True)
    if hasattr(sys.stderr, 'reconfigure'):
        sys.stderr.reconfigure(encoding='utf-8', write_through=True)
except Exception:
    pass

# ============ CONFIG & CONSTANTS ============
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')

SPREADSHEET_ID = "1B-QCbEnPpILFFEWPYheGdmkgYV9gSf4lAyQMlhzwOCM"
GTALK_OA_TOKEN = os.environ.get("OPR_GTALK_OA_TOKEN") or "2077276776281051136:8hMHvBBU8qXKps3mLPzgKBucPLSQPg3Y"

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

SERVICE_ACCOUNT_CANDIDATES = [
    os.path.join(BASE_DIR, 'credentials.json'),
    r'C:\Users\lap4all\Documents\Auto report\credentials.json',
    r'C:\Users\lap4all\Desktop\Backlog_Automation\credentials.json',
    r'C:\Users\lap4all\Downloads\credentials.json',
    'credentials.json',
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



def normalize_str(s):
    if not s:
        return ""
    return unicodedata.normalize("NFC", str(s)).strip()


def normalize_am_name(name):
    if not name or pd.isna(name):
        return "Không xác định"
    clean = " ".join(str(name).strip().split())
    if clean == "" or clean.lower() == "#n/a" or clean.lower() == "nan":
        return "Không xác định"
    
    prefix = ""
    name_part = clean
    dash_idx = clean.find("-")
    if dash_idx != -1:
        prefix_part = clean[:dash_idx]
        if prefix_part.isdigit():
            prefix = clean[:dash_idx + 1]
            name_part = clean[dash_idx + 1:]
            
    words = name_part.split(' ')
    normalized_words = []
    for w in words:
        if w.upper() == "AM":
            normalized_words.append("AM")
        elif len(w) > 0:
            normalized_words.append(w[0].upper() + w[1:].lower())
        else:
            normalized_words.append("")
    return prefix + " ".join(normalized_words)


DEFAULT_AM_GTALK_MAPPING = {
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
    "Cao Thị Thanh Thủy": "2083241927281995776"
}


def get_am_channel_map():
    mapping = dict(DEFAULT_AM_GTALK_MAPPING)
    # Check if am_channel_map.json exists to supplement channel IDs
    map_json_path = os.path.join(BASE_DIR, "am_channel_map.json")
    if os.path.exists(map_json_path):
        try:
            with open(map_json_path, "r", encoding="utf-8-sig") as f:
                json_map = json.load(f)
                if isinstance(json_map, dict):
                    for k, v in json_map.items():
                        if k and v:
                            mapping[normalize_str(k)] = str(v).strip()
        except Exception as e:
            print(f"⚠️ Không đọc được file am_channel_map.json: {e}")
    return mapping


def find_channel_id_for_am(am_name, am_map):
    norm_target = normalize_str(am_name).lower()
    # Remove 'am ' prefix if present
    if norm_target.startswith("am "):
        norm_target = norm_target[3:].strip()
        
    for k, channel_id in am_map.items():
        norm_k = normalize_str(k).lower()
        if norm_k.startswith("am "):
            norm_k = norm_k[3:].strip()
            
        if norm_target == norm_k or norm_k in norm_target or norm_target in norm_k:
            return channel_id
            
    return None


def run_pipeline(dry_run=False):
    now_str = datetime.now().strftime('%d/%m/%Y %H:%M')
    print(f"🚀 BẮT ĐẦU GỬI CẢNH BÁO RIÊNG TỪNG AM (TTS PICKUP) LÚC: {now_str}")
    
    # 1. Authorize Google Sheet
    try:
        gc_client = get_gspread_client(spreadsheet_id=SPREADSHEET_ID)
        sh = gc_client.open_by_key(SPREADSHEET_ID)
        print(f"✔️ Đã kết nối thành công tới: '{sh.title}'")
    except Exception as e:
        print(f"❌ Lỗi kết nối Google Sheets: {e}")
        sys.exit(1)
        
    # 2. Read Worksheets
    try:
        ws_raw = sh.worksheet("raw")
        raw_values = ws_raw.get_all_records()
        df_raw = pd.DataFrame(raw_values)
        print(f"✔️ Đã đọc dữ liệu tab 'raw' ({len(df_raw)} dòng)")
    except Exception as e:
        print(f"❌ Lỗi đọc dữ liệu tab 'raw': {e}")
        sys.exit(1)
        
    try:
        ws_cocau = sh.worksheet("CoCauVung")
        cocau_data = ws_cocau.get_all_values()
        cocau_map = {}
        for row in cocau_data[1:]:
            if len(row) >= 4:
                w_id = str(row[0]).strip()
                bc_name = str(row[1]).strip()
                tinh = str(row[2]).strip()
                am = str(row[3]).strip()
                if w_id:
                    cocau_map[w_id] = {
                        "Bưu cục": bc_name,
                        "Tỉnh": tinh,
                        "AM": am
                    }
        print(f"✔️ Đã đọc dữ liệu tab 'CoCauVung' ({len(cocau_map)} mã kho)")
    except Exception as e:
        print(f"❌ Lỗi đọc dữ liệu tab 'CoCauVung': {e}")
        sys.exit(1)

    if df_raw.empty:
        print("⚠️ Tab 'raw' trống. Không có dữ liệu để xử lý.")
        sys.exit(0)

    # Find sheet links for Master tab and per-AM tabs
    all_worksheets = {ws.title: ws.id for ws in sh.worksheets()}
    master_tab_name = "Đơn Lấy TTS Chưa Gán"
    if master_tab_name in all_worksheets:
        master_sheet_link = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit#gid={all_worksheets[master_tab_name]}"
    else:
        master_sheet_link = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit"

    am_sheet_links = {}
    for title, sheet_id in all_worksheets.items():
        if title.startswith("[TTS Lấy]"):
            am_clean_name = title.replace("[TTS Lấy]", "").strip()
            am_sheet_links[am_clean_name] = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit#gid={sheet_id}"

    # 3. Filter for: Loại đơn = "Lấy", Khách hàng = "TTS"
    df_raw['Loại đơn normalized'] = df_raw['Loại đơn'].astype(str).str.strip().str.lower()
    df_raw['Khách hàng normalized'] = df_raw['Khách hàng'].astype(str).str.strip().str.lower()
    
    df_tts_pickup = df_raw[
        (df_raw['Loại đơn normalized'] == 'lấy') & 
        (df_raw['Khách hàng normalized'] == 'tts')
    ]
    
    total_tts_pickup = len(df_tts_pickup)
    print(f"🎯 Tổng số đơn lấy TTS phát sinh: {total_tts_pickup}")
    
    if total_tts_pickup == 0:
        print("🎉 Không có đơn lấy TTS nào phát sinh hôm nay!")
        sys.exit(0)

    # 4. Process orders per AM
    am_stats = {}
    am_unassigned_bcs = {} # am_name -> { buucuc_name: [order_ids] }

    for idx, row in df_tts_pickup.iterrows():
        poc_raw = str(row['Mã bưu cục']).strip()
        order_id = str(row.get('Mã đơn hàng', '')).strip()
        
        po_info = None
        if poc_raw in cocau_map:
            po_info = cocau_map[poc_raw]
        else:
            for cid, info in cocau_map.items():
                if cid.startswith(poc_raw) or poc_raw == cid:
                    po_info = info
                    break
                    
        if po_info:
            am_name = normalize_am_name(po_info["AM"])
            buucuc_name = po_info["Bưu cục"]
        else:
            am_name = "Không xác định"
            buucuc_name = f"Bưu cục {poc_raw}"
            
        status = str(row['Trạng thái']).strip()
        is_unassigned = (status == "Chưa có chuyến đi trong ngày")
        is_assigned = status in ["Đang có chuyến đi trong ngày", "Đã có chuyến đi trong ngày"]
        
        if am_name not in am_stats:
            am_stats[am_name] = {"total": 0, "assigned": 0, "unassigned": 0}
            
        am_stats[am_name]["total"] += 1
        
        if is_assigned:
            am_stats[am_name]["assigned"] += 1
        elif is_unassigned:
            am_stats[am_name]["unassigned"] += 1
            
            if am_name not in am_unassigned_bcs:
                am_unassigned_bcs[am_name] = {}
            if buucuc_name not in am_unassigned_bcs[am_name]:
                am_unassigned_bcs[am_name][buucuc_name] = []
            if order_id:
                am_unassigned_bcs[am_name][buucuc_name].append(order_id)

    # Regional Totals
    grand_total = sum(s["total"] for s in am_stats.values())
    grand_assigned = sum(s["assigned"] for s in am_stats.values())
    grand_unassigned = sum(s["unassigned"] for s in am_stats.values())
    grand_rate = (grand_assigned / grand_total * 100.0) if grand_total > 0 else 0.0

    print(f"📊 Toàn vùng: Tổng phát sinh = {grand_total} | Đã gán = {grand_assigned} | Chưa gán = {grand_unassigned} ({grand_rate:.1f}%)")

    # Filter AMs that have pending (unassigned) orders
    pending_ams = [am for am, bcs in am_unassigned_bcs.items() if sum(len(orders) for orders in bcs.values()) > 0]
    
    # Sort AMs by unassigned count descending to determine ranking
    # Rank #1 = AM with highest number of unassigned orders
    sorted_pending_ams = sorted(
        pending_ams,
        key=lambda am: (sum(len(orders) for orders in am_unassigned_bcs[am].values()), am_stats[am]["total"]),
        reverse=True
    )

    total_pending_ams_count = len(sorted_pending_ams)
    print(f"🚨 Có {total_pending_ams_count} AM có đơn TTS pending chưa gán chuyến đi.")

    am_channel_map = get_am_channel_map()

    # 5. Format & Send message for each AM
    for rank_idx, am_name in enumerate(sorted_pending_ams, 1):
        stats = am_stats[am_name]
        tot = stats["total"]
        ass = stats["assigned"]
        unass = stats["unassigned"]
        rate = (ass / tot * 100.0) if tot > 0 else 0.0
        
        bcs_dict = am_unassigned_bcs[am_name]
        
        # Build detail per Post Office
        buucuc_details_html = ""
        for bc_name, order_list in sorted(bcs_dict.items(), key=lambda x: len(x[1]), reverse=True):
            count_bc = len(order_list)
            bc_name_clean = html.escape(str(bc_name), quote=False)
            # Format MVĐs as code tags, limit display to max 20 MVĐs per post office to avoid payload overflow
            max_mvds = 20
            if len(order_list) > max_mvds:
                shown_mvds = order_list[:max_mvds]
                rem = len(order_list) - max_mvds
                mvd_formatted = ", ".join([f"<code>{html.escape(str(mvd), quote=False)}</code>" for mvd in shown_mvds]) + f" <i>(và {rem} đơn khác...)</i>"
            else:
                mvd_formatted = ", ".join([f"<code>{html.escape(str(mvd), quote=False)}</code>" for mvd in order_list]) if order_list else "<i>(Không có mã)</i>"
            buucuc_details_html += f"  • 🏢 <b>{bc_name_clean}</b> ({count_bc} đơn):\n    - MVĐ: {mvd_formatted}\n"

        am_link = am_sheet_links.get(am_name, master_sheet_link)

        msg = f"🚨 <b>ĐƠN LẤY TTS CHƯA GÁN CHUYẾN ĐI LẤY HÀNG</b>\n"
        msg += f"⏱️ <b>Mốc cập nhật:</b> {now_str}\n"
        msg += f"🔗 <b>Tổng hợp toàn vùng:</b> <a href=\"{master_sheet_link}\"><b>Xem chi tiết</b></a>\n\n"
        
        msg += f"👤 <b>ĐƠN PENDING TTS CHƯA GÁN CHUYẾN LẤY HÀNG - {html.escape(am_name.upper(), quote=False)}</b>\n"
        msg += f"  • 📥 <b>Tổng đơn phát sinh của AM:</b> <b>{tot}</b> đơn\n"
        msg += f"  • ✅ <b>Đã gán:</b> <b>{ass}</b> đơn ({rate:.1f}%)\n"
        msg += f"  • ❌ <b>Chưa gán (Pending):</b> <b>{unass}</b> đơn\n"
        msg += f"  • 🏆 <b>Xếp hạng pending:</b> Hạng <b>#{rank_idx}</b> / {total_pending_ams_count} AM có lượng đơn pending cao nhất\n\n"
        
        msg += f"📍 <b>DANH SÁCH BƯU CỤC & MÃ VẬN ĐƠN (MVĐ) CHƯA GÁN:</b>\n"
        msg += buucuc_details_html + "\n"
        
        msg += f"👉 Anh/chị AM <b>{html.escape(am_name, quote=False)}</b> vui lòng push bưu cục gán chuyến đi lấy hàng ngay trước <b>12:00</b> nhé!\n"
        msg += f"🔗 <a href=\"{am_link}\"><b>Ấn vào đây để mở Sheet chi tiết đơn của AM {html.escape(am_name, quote=False)}</b></a>"

        channel_id = find_channel_id_for_am(am_name, am_channel_map)
        
        print("==================================================")
        print(f"📩 TIN NHẮN DÀNH CHO AM: {am_name} (Hạng #{rank_idx} | Channel ID: {channel_id or 'CHƯA CÓ'})")
        print("--------------------------------------------------")
        print(msg)
        print("==================================================")

        if dry_run:
            print(f"🔍 [DRY RUN] Đã tạo tin nhắn cho {am_name}. Bỏ qua gửi GTalk.\n")
            continue

        if not channel_id:
            print(f"⚠️ Không tìm thấy Channel ID cho AM '{am_name}'. Bỏ qua gửi GTalk.\n")
            continue

        # Send via GTalk API
        url_send = "https://mbff.ghn.vn/api/gtalk/send-message"
        client_msg_id = str(int(time.time() * 1000))
        payload = {
            "channelId": channel_id,
            "clientMsgId": client_msg_id,
            "content": {
                "parseMode": "HTML",
                "text": msg
            },
            "oaToken": GTALK_OA_TOKEN
        }
        
        try:
            res = requests.post(url_send, json=payload, headers={"Content-Type": "application/json"}, timeout=20, verify=False)
            if res.status_code == 200:
                res_data = res.json()
                if res_data.get("errorCode") == "success":
                    print(f"✅ Đã gửi thành công cho AM {am_name} (Channel: {channel_id})")
                else:
                    print(f"❌ Gửi cho {am_name} lỗi API: {res_data.get('error')}")
            else:
                print(f"❌ Gửi cho {am_name} lỗi HTTP {res.status_code}: {res.text}")
        except Exception as e:
            print(f"❌ Lỗi kết nối gửi tin nhắn cho AM {am_name}: {e}")
            
        time.sleep(1) # Small pause between requests

    print("🎉 HOÀN THÀNH QUY TRÌNH GỬI CẢNH BÁO TTS PICKUP RIÊNG TỪNG AM!")


def main():
    parser = argparse.ArgumentParser(description="Push individual TTS pickup alerts to AMs via GTalk.")
    parser.add_argument("--dry-run", action="store_true", help="Print messages without sending GTalk requests.")
    args = parser.parse_args()

    try:
        run_pipeline(dry_run=args.dry_run)
    except Exception as e:
        import traceback
        now_str = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
        tb_str = traceback.format_exc()
        error_msg = f"⚠️ <b>[BOT LỖI - PUSH AM TTS PICKUP INDIVIDUAL]</b>\n" \
                    f"⏱️ <b>Thời gian:</b> {now_str}\n" \
                    f"❌ <b>Lỗi ngắn:</b> <code>{str(e)}</code>\n\n" \
                    f"🛠️ <b>Traceback chi tiết:</b>\n<pre>{tb_str[:1500]}</pre>"
        
        print("❌ PHÁT HIỆN LỖI CRASH SCRIPT. ĐANG GỬI TIN BÁO LỖI NỘI BỘ...")
        print(tb_str)
        
        env_path = r"c:\Users\lap4all\Desktop\New folder\.env"
        if os.path.exists(env_path):
            load_dotenv(dotenv_path=env_path, override=True)
        else:
            load_dotenv()
            
        gtalk_admin = os.environ.get("OPR_GTALK_CHANNEL_ID_ADMIN") or os.environ.get("GTALK_CHANNEL_ID_ADMIN") or "2067164759710552066"
        
        if gtalk_admin:
            try:
                url_gtalk = "https://mbff.ghn.vn/api/gtalk/send-message"
                payload_gtalk = {
                    "channelId": gtalk_admin,
                    "clientMsgId": str(int(time.time() * 1000)),
                    "content": {
                        "parseMode": "HTML",
                        "text": error_msg
                    },
                    "oaToken": GTALK_OA_TOKEN
                }
                requests.post(url_gtalk, json=payload_gtalk, headers={"Content-Type": "application/json"}, timeout=10, verify=False)
                print("🟢 Đã gửi báo lỗi qua GTalk nội bộ.")
            except Exception as gtalk_err:
                print(f"❌ Không gửi được báo lỗi qua GTalk: {gtalk_err}")

        sys.exit(1)


if __name__ == "__main__":
    main()
