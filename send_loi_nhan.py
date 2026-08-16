# -*- coding: utf-8 -*-
"""
Script: send_loi_nhan.py
Đọc nội dung lời nhắn tại Cột H và danh sách ID Group ở Cột A/B từ tab 'lời nhắn'
trên Google Sheet, sau đó gửi tin nhắn hàng loạt qua GTalk API.

Cách dùng:
- Dry run (chỉ kiểm tra & xem trước tất cả):
    python send_loi_nhan.py
- Gửi thật tất cả:
    python send_loi_nhan.py --send
- Lọc theo AM (chỉ xem trước):
    python send_loi_nhan.py --am "Nguyễn Văn A" "Trần Văn B"
- Lọc theo AM và gửi thật:
    python send_loi_nhan.py --am "Nguyễn Văn A" --send
"""

import os
import sys
import time
import json
import argparse
import requests
import gspread
from google.oauth2.service_account import Credentials
from google.oauth2.credentials import Credentials as UserCredentials

import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Mã hóa UTF-8 cho console Windows
os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    if hasattr(sys.stderr, 'reconfigure'):
        sys.stderr.reconfigure(encoding='utf-8')
except Exception:
    pass

SPREADSHEET_ID = "1MtbZBgRFwCWj6uQKsSqddiJ2GsTiEvKxRIPSshDa5PM"
TAB_NAME = "lời nhắn"
DEFAULT_TOKEN = "2077276776281051136:8hMHvBBU8qXKps3mLPzgKBucPLSQPg3Y"
GTALK_API_URL = "https://mbff.ghn.vn/api/gtalk/send-message"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

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
    "authorized_user.json",
]

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


def get_gspread_client(spreadsheet_id=None):
    """Thử Service Account trước, fallback sang authorized_user.json / OAuth nếu bị từ chối quyền."""
    for cred_path in SERVICE_ACCOUNT_CANDIDATES:
        if os.path.isfile(cred_path):
            try:
                creds = Credentials.from_service_account_file(cred_path, scopes=SCOPES)
                gc = gspread.authorize(creds)
                if spreadsheet_id:
                    gc.open_by_key(spreadsheet_id)
                return gc
            except Exception:
                pass

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

    oauth_file = os.path.join(BASE_DIR, "credentials_oauth.json")
    if os.path.exists(oauth_file):
        try:
            gc = gspread.oauth(
                credentials_filename=oauth_file,
                authorized_user_filename=os.path.join(BASE_DIR, "authorized_user.json")
            )
            if spreadsheet_id:
                gc.open_by_key(spreadsheet_id)
            return gc
        except Exception:
            pass

    raise PermissionError("Không thể xác thực Google Sheets bằng credentials.json hoặc authorized_user.json")


def load_sheet_data():
    gc = get_gspread_client(SPREADSHEET_ID)
    sh = gc.open_by_key(SPREADSHEET_ID)
    
    # Tìm tab 'lời nhắn' hoặc tab có id 1525919864
    ws = None
    for item in sh.worksheets():
        if str(item.id) == "1525919864" or "lời" in item.title.lower():
            ws = item
            break
    if not ws:
        ws = sh.worksheet(TAB_NAME)
        
    return ws.get_all_values()


def parse_loinhan_data(rows):
    token = DEFAULT_TOKEN
    common_msg = ""
    target_groups = []

    # Check row 2 col B for token
    if len(rows) >= 2 and len(rows[1]) > 1 and rows[1][1].strip():
        val = rows[1][1].strip()
        if ":" in val:
            token = val

    # Check row 1 col H (index 7) for common message
    if len(rows) >= 1 and len(rows[0]) > 7 and rows[0][7].strip():
        common_msg = rows[0][7].strip()

    # Parse rows starting from row 3 (index 2)
    for i in range(2, len(rows)):
        row = rows[i]
        am_name = row[0].strip() if len(row) > 0 else ""
        group_id = row[1].strip() if len(row) > 1 else ""
        row_msg = row[7].strip() if len(row) > 7 else ""

        # Dùng row_msg nếu có, không thì dùng common_msg
        msg_to_send = row_msg if row_msg else common_msg

        if am_name and group_id and group_id.isdigit():
            target_groups.append({
                "row_index": i + 1,
                "am_name": am_name,
                "group_id": group_id,
                "message": msg_to_send
            })

    return token, target_groups


def send_gtalk_message(group_id, text, oa_token):
    payload = {
        "channelId": str(group_id),
        "clientMsgId": str(int(time.time() * 1000)),
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
    parser = argparse.ArgumentParser(description="Gửi lời nhắn hàng loạt qua GTalk")
    parser.add_argument("--send", "--force-send", action="store_true", help="Thực hiện gửi thật qua GTalk API")
    parser.add_argument("--am", nargs="+", help="Lọc chỉ gửi cho các AM chỉ định (tên hoặc 1 phần tên). VD: --am 'Văn A' 'Trần B'")
    args = parser.parse_args()

    is_dry_run = not args.send

    print("=" * 65)
    print("🚀 CHƯƠNG TRÌNH GỬI LỜI NHẮN HÀNG LOẠT THEO GROUP ID")
    print(f"📌 Chế độ: {'DRY-RUN (Chỉ xem trước, KHÔNG gửi thật)' if is_dry_run else '🔴 KHỞI CHẠY GỬI THẬT'}")
    print("=" * 65)

    print("📥 Đang đọc dữ liệu từ Google Sheet...")
    rows = load_sheet_data()
    token, target_groups = parse_loinhan_data(rows)

    print(f"🔑 Token GTalk OA: {token[:20]}...")
    print(f"📋 Tìm thấy tổng cộng {len(target_groups)} group ID từ Sheet.")

    # Lọc theo tham số --am nếu có
    filter_ams = []
    if args.am:
        for arg in args.am:
            for sub in arg.split(","):
                if sub.strip():
                    filter_ams.append(sub.strip().lower())

    if filter_ams:
        filtered = [
            item for item in target_groups
            if any(f_am in item["am_name"].lower() for f_am in filter_ams)
        ]
        print(f"🎯 Đã lọc {len(filtered)}/{len(target_groups)} group thuộc AM khớp với tìm kiếm: {', '.join(args.am)}")
        target_groups = filtered

    print()

    if not target_groups:
        print("⚠️ Không tìm thấy Group ID nào phù hợp!")
        return

    # In mẫu lời nhắn
    sample_msg = target_groups[0]['message']
    print("💬 NỘI DUNG LỜI NHẮN SẼ GỬI:")
    print("-" * 50)
    print(sample_msg)
    print("-" * 50)

    success_count = 0
    fail_count = 0

    print(f"\nDanh sách {len(target_groups)} AM & Group ID:")
    for idx, item in enumerate(target_groups, 1):
        am = item["am_name"]
        gid = item["group_id"]
        msg = item["message"]

        print(f" [{idx:02d}/{len(target_groups):02d}] AM: {am:<25} | Group ID: {gid}", end="")

        if is_dry_run:
            print(" -> [DRY-RUN (Chưa gửi)]")
            success_count += 1
        else:
            ok, err = send_gtalk_message(gid, msg, token)
            if ok:
                print(" -> ✅ Thành công")
                success_count += 1
            else:
                print(f" -> ❌ Thất bại: {err}")
                fail_count += 1
            time.sleep(0.3)

    print("\n" + "=" * 65)
    print("📊 TỔNG KẾT XỬ LÝ:")
    print(f"   - Chế độ:    {'DRY-RUN' if is_dry_run else 'THỰC HÀNH GỬI THẬT'}")
    print(f"   - Thành công: {success_count}/{len(target_groups)}")
    if not is_dry_run:
        print(f"   - Thất bại:   {fail_count}/{len(target_groups)}")
    else:
        print("💡 Chạy lệnh 'python send_loi_nhan.py --send' để gửi thật!")
    print("=" * 65)


if __name__ == "__main__":
    main()
