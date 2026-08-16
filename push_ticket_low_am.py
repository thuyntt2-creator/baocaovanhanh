# -*- coding: utf-8 -*-
"""
Script: push_ticket_low_am.py
Đọc sheet 'ticket' từ Google Sheet, lọc thông tin từ phần:
  '━━ AM tồn thấp (≤10 phiếu) — tổng quan ━━'
lấy số phiếu tồn từng AM và gửi thông báo qua GTalk tới ID group tương ứng ở cột D.

Cách dùng:
- Xem trước (Dry Run, không gửi thật):
    python push_ticket_low_am.py
- Gửi thật:
    python push_ticket_low_am.py --send
"""

import os
import sys
import re
import time
import json
import argparse
import unicodedata
from datetime import datetime
import requests
import gspread
from google.oauth2.service_account import Credentials
from google.oauth2.credentials import Credentials as UserCredentials

import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Cấu hình mã hóa UTF-8 cho console Windows
os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    if hasattr(sys.stderr, 'reconfigure'):
        sys.stderr.reconfigure(encoding='utf-8')
except Exception:
    pass

SPREADSHEET_ID = "1MtbZBgRFwCWj6uQKsSqddiJ2GsTiEvKxRIPSshDa5PM"
TAB_NAME = "ticket"
DEFAULT_TOKEN = "2077276776281051136:8hMHvBBU8qXKps3mLPzgKBucPLSQPg3Y"
GTALK_API_URL = "https://mbff.ghn.vn/api/gtalk/send-message"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

SERVICE_ACCOUNT_CANDIDATES = [
    os.path.join(BASE_DIR, "credentials.json"),
    r"C:\Users\lap4all\Documents\Auto report\credentials.json",
    r"C:\Users\lap4all\Desktop\Backlog_Automation\credentials.json",
    "credentials.json",
]

AUTH_USER_CANDIDATES = [
    os.path.join(BASE_DIR, "authorized_user.json"),
    r"C:\Users\lap4all\Documents\Auto report\authorized_user.json",
    r"C:\Users\lap4all\Desktop\Backlog_Automation\authorized_user.json",
    "authorized_user.json"
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


def norm(s):
    if not s:
        return ""
    return unicodedata.normalize("NFC", str(s)).strip().lower()


def load_sheet_data():
    gc = get_gspread_client(spreadsheet_id=SPREADSHEET_ID)
    sh = gc.open_by_key(SPREADSHEET_ID)
    ws = sh.worksheet(TAB_NAME)
    return ws.get_all_values()


def parse_am_group_mapping(rows):
    """
    Đọc mapping AM -> Group ID từ Cột C và D.
    Cột C: Tên AM
    Cột D: Group ID
    Row 2 Col D: Token (nếu có)
    """
    am_map = {}
    token = DEFAULT_TOKEN

    for i, row in enumerate(rows):
        c_val = str(row[2]).strip() if len(row) > 2 else ""
        d_val = str(row[3]).strip() if len(row) > 3 else ""

        if c_val == "AM" and d_val:
            token = d_val
        elif c_val and d_val and c_val != "Email_AM":
            am_map[norm(c_val)] = {
                "display_name": c_val,
                "group_id": d_val
            }

    return am_map, token


def parse_low_ticket_section(rows):
    """
    Lọc danh sách các AM thuộc mục '━━ AM tồn thấp (≤10 phiếu) — tổng quan ━━'
    """
    in_section = False
    items = []

    for row in rows:
        col_a = str(row[0]).strip() if len(row) > 0 else ""
        if "AM tồn thấp" in col_a and "tổng quan" in col_a:
            in_section = True
            continue
        if in_section:
            if col_a.startswith("━━") or (col_a.startswith("(") and "Auto GLT" in col_a):
                break
            if col_a.startswith("•"):
                items.append(col_a)

    return items


def extract_am_name_from_line(line):
    """
    Ví dụ line: '• Huỳnh Thị Kim Chi (mã 3062894): 6 phiếu / 2 bưu cục'
    Trả về: 'Huỳnh Thị Kim Chi'
    """
    m = re.search(r"•\s*(.*?)\s*\(\s*mã\s*\d+\s*\)", line)
    if m:
        return m.group(1).strip()
    return None


def send_gtalk_message(group_id, text, oa_token):
    payload = {
        "channelId": str(group_id),
        "clientMsgId": str(int(datetime.now().timestamp() * 1000)),
        "content": {
            "parseMode": "HTML",
            "text": text
        },
        "oaToken": oa_token
    }
    try:
        r = requests.post(GTALK_API_URL, json=payload, timeout=15, verify=False)
        try:
            res = r.json()
        except Exception:
            res = {}

        if r.status_code == 200 and res.get("errorCode") == "success":
            return True, "Thành công"
        return False, f"HTTP {r.status_code} — {r.text}"
    except Exception as e:
        return False, str(e)


def main():
    parser = argparse.ArgumentParser(description="Gửi cảnh báo Ticket tồn cho AMs")
    parser.add_argument("--send", "--force-send", action="store_true", help="Thực hiện gửi thật qua GTalk API")
    args = parser.parse_args()

    is_dry_run = not args.send

    print("=" * 60)
    print("🚀 BẮT ĐẦU XỬ LÝ CẢNH BÁO TICKET TỒN AM")
    print(f"📌 Chế độ: {'DRY-RUN (Chỉ xem trước, KHÔNG gửi thật)' if is_dry_run else '🔴 KHỞI CHẠY GỬI THẬT'}")
    print("=" * 60)

    print("📥 Đang tải dữ liệu từ Google Sheet...")
    rows = load_sheet_data()
    print(f"✅ Đã tải {len(rows)} dòng từ tab '{TAB_NAME}'.")

    am_map, token = parse_am_group_mapping(rows)
    print(f"🔑 Token GTalk OA: {token[:15]}...")
    print(f"📋 Đã nạp {len(am_map)} AM từ cột C/D.")

    ticket_lines = parse_low_ticket_section(rows)
    print(f"🎯 Tìm thấy {len(ticket_lines)} phiếu tồn trong mục 'AM tồn thấp (≤10 phiếu)'.\n")

    success_count = 0
    fail_count = 0
    skip_count = 0

    for line in ticket_lines:
        am_name = extract_am_name_from_line(line)
        if not am_name:
            print(f"⚠️ Không trích xuất được tên AM từ dòng: {line}")
            skip_count += 1
            continue

        am_info = am_map.get(norm(am_name))

        # Format nội dung gửi
        msg_text = f"TỒN PHIẾU HỐI GIAO/LẤY/TRẢ\n{line}"

        if not am_info:
            print(f"❌ AM '{am_name}' không tìm thấy ID Group ở cột C/D! (Bỏ qua)")
            print(f"   Nội dung dự kiến: \n{msg_text}\n")
            skip_count += 1
            continue

        group_id = am_info["group_id"]
        display_name = am_info["display_name"]

        print(f"👤 AM: {display_name}")
        print(f"🆔 Group ID: {group_id}")
        print("💬 Nội dung gửi:")
        print("-" * 40)
        print(msg_text)
        print("-" * 40)

        if is_dry_run:
            print("🔍 [DRY-RUN] Sẵn sàng gửi. (Dùng flag --send để gửi thật)\n")
            success_count += 1
        else:
            ok, err = send_gtalk_message(group_id, msg_text, token)
            if ok:
                print("✅ Gửi tin nhắn thành công!\n")
                success_count += 1
            else:
                print(f"❌ Gửi thất bại: {err}\n")
                fail_count += 1
            time.sleep(0.5)

    print("=" * 60)
    print("📊 TỔNG KẾT XỬ LÝ:")
    print(f"   - Thành công: {success_count}")
    print(f"   - Thất bại:   {fail_count}")
    print(f"   - Bỏ qua:     {skip_count}")
    print("=" * 60)


if __name__ == "__main__":
    main()
