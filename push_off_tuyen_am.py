# -*- coding: utf-8 -*-
"""
Script: push_off_tuyen_am.py
Đọc kết quả Off tuyến từ tab 'Đang OFF' thuộc Google Sheet:
https://docs.google.com/spreadsheets/d/1PjzFqJO-wkQ8SNsPHD721_CbPr6c_ArZKuGGU6KqDZg

Tính năng:
1. Gom nhóm dữ liệu theo BƯU CỤC & AM, liệt kê rõ từng Xã/Phường tắt theo từng Bưu cục.
2. Gửi tin nhắn văn bản đã định dạng rõ ràng (kèm link Google Sheet) tới Group ID riêng của từng AM qua GTalk API.
3. Không gửi ảnh (chỉ gửi tin nhắn văn bản gọn nhẹ).

Cách dùng:
- Dry Run (chỉ xem trước tin nhắn, không gửi GTalk):
    python push_off_tuyen_am.py
- Gửi thật qua GTalk API:
    python push_off_tuyen_am.py --send
"""

import os
import sys
import time
import json
import argparse
import unicodedata
import requests
import gspread
from google.oauth2.service_account import Credentials

import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Mã hóa UTF-8 cho Windows Console
os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    if hasattr(sys.stderr, 'reconfigure'):
        sys.stderr.reconfigure(encoding='utf-8')
except Exception:
    pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

SPREADSHEET_ID = "1PjzFqJO-wkQ8SNsPHD721_CbPr6c_ArZKuGGU6KqDZg"
SHEET_LINK = "https://docs.google.com/spreadsheets/d/1PjzFqJO-wkQ8SNsPHD721_CbPr6c_ArZKuGGU6KqDZg/edit?gid=1524249564#gid=1524249564"
TAB_NAME = "Đang OFF"

# Token GTalk OA mới
DEFAULT_TOKEN = "2077276776281051136:8hMHvBBU8qXKps3mLPzgKBucPLSQPg3Y"
GTALK_API_URL = "https://mbff.ghn.vn/api/gtalk/send-message"

SERVICE_ACCOUNT_CANDIDATES = [
    os.path.join(BASE_DIR, "credentials.json"),
    r"C:\Users\lap4all\Desktop\Backlog_Automation\credentials.json",
    r"C:\Users\lap4all\Downloads\credentials.json",
    "credentials.json",
]

# Mapping Tên AM -> ID Group riêng biệt của AM
AM_GROUP_MAP = {
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
    "Lê Hồng Minh Tâm": "2079827054540226560"
}


def normalize_str(s):
    if not s:
        return ""
    return unicodedata.normalize("NFC", str(s)).strip()


def get_service_account():
    for candidate in SERVICE_ACCOUNT_CANDIDATES:
        if os.path.exists(candidate):
            return candidate
    raise FileNotFoundError("❌ Không tìm thấy file credentials.json.")


def load_sheet_data():
    cred_file = get_service_account()
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_file(cred_file, scopes=scopes)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SPREADSHEET_ID)

    ws = None
    for item in sh.worksheets():
        if str(item.id) == "1524249564" or "off" in item.title.lower():
            ws = item
            break
    if not ws:
        ws = sh.worksheet(TAB_NAME)

    return ws.get_all_values()


def parse_off_data(rows):
    if not rows or len(rows) < 2:
        return {}

    norm_am_map = {normalize_str(k).lower(): k for k in AM_GROUP_MAP.keys()}

    am_data = {}

    for i in range(1, len(rows)):
        row = rows[i]
        if not any(row):
            continue

        province = normalize_str(row[0]) if len(row) > 0 else ""
        district = normalize_str(row[1]) if len(row) > 1 else ""
        ward = normalize_str(row[2]) if len(row) > 2 else ""
        ward_id = normalize_str(row[3]) if len(row) > 3 else ""
        post_office = normalize_str(row[4]) if len(row) > 4 else ""
        hrbp_confirm = normalize_str(row[5]) if len(row) > 5 else ""
        result = normalize_str(row[6]) if len(row) > 6 else ""
        cap_down = normalize_str(row[7]) if len(row) > 7 else ""
        off_from = normalize_str(row[8]) if len(row) > 8 else ""
        off_to = normalize_str(row[9]) if len(row) > 9 else ""
        am_name = normalize_str(row[10]) if len(row) > 10 else ""

        if not am_name:
            continue

        cap_down_fmt = cap_down
        try:
            val = float(cap_down)
            if val <= 1:
                cap_down_fmt = f"{int(val * 100)}%"
        except ValueError:
            pass

        norm_key = am_name.lower()
        canonical_am = norm_am_map.get(norm_key, am_name)

        ward_info = {
            "province": province,
            "district": district,
            "ward": ward,
            "post_office": post_office,
            "result": result or "DUYỆT",
            "cap_down": cap_down_fmt,
            "off_from": off_from,
            "off_to": off_to,
            "am": canonical_am
        }

        bc_key = post_office or "CHƯA XÁC ĐỊNH"

        # Gom theo AM -> Bưu cục
        if canonical_am not in am_data:
            am_data[canonical_am] = {}
        if bc_key not in am_data[canonical_am]:
            am_data[canonical_am][bc_key] = {
                "bc_name": bc_key,
                "result": result or "DUYỆT",
                "cap_down": cap_down_fmt,
                "off_from": off_from,
                "off_to": off_to,
                "wards": []
            }
        am_data[canonical_am][bc_key]["wards"].append(ward_info)

    return am_data


def format_am_text_message(am_name, am_bcs):
    """
    Định dạng tin nhắn riêng cho từng AM gom nhóm theo Bưu cục và liệt kê các xã/phường + Link Google Sheet.
    """
    total_wards = sum(len(b["wards"]) for b in am_bcs.values())
    total_bcs = len(am_bcs)

    msg = f"📢 <b>THÔNG BÁO KẾT QUẢ OFF TUYẾN — AM: {am_name}</b>\n"
    msg += f"📊 <b>Tổng số:</b> {total_wards} Xã/Phường | {total_bcs} Bưu cục\n"

    for bc_name, b_info in am_bcs.items():
        msg += f"\n📦 <b>BƯU CỤC: {bc_name}</b> ({len(b_info['wards'])} xã/phường)\n"
        msg += f"   • Trạng thái: <b>{b_info['result']}</b> (Cap Down: {b_info['cap_down']})\n"
        if b_info['off_from'] or b_info['off_to']:
            msg += f"   • Thời gian: {b_info['off_from']} ➔ {b_info['off_to']}\n"
        msg += "   📍 <i>Các xã/phường off:</i>\n"

        for idx, w in enumerate(b_info['wards'], 1):
            msg += f"      {idx}. <b>{w['ward']}</b> - {w['district']}, {w['province']}\n"

    msg += f"\n🔗 <b>Xem chi tiết tại Sheet:</b> <a href=\"{SHEET_LINK}\">Tab Đang OFF</a>"
    return msg.strip()


def send_gtalk_text_message(group_id, text, oa_token):
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
    parser = argparse.ArgumentParser(description="Gửi thông báo Off tuyến văn bản cho AM")
    parser.add_argument("--send", "--force-send", action="store_true", help="Gửi thật qua GTalk API")
    args = parser.parse_args()

    is_dry_run = not args.send

    print("=" * 65)
    print("🚀 CHƯƠNG TRÌNH THỐNG KÊ VÀ BÁO CÁO KẾT QUẢ OFF TUYẾN (TEXT ONLY)")
    print(f"🔑 Token GTalk OA: {DEFAULT_TOKEN[:20]}...")
    print(f"📌 Chế độ: {'DRY-RUN (Chỉ xem trước, KHÔNG gửi thật)' if is_dry_run else '🔴 KHỞI CHẠY GỬI THẬT QUA GTALK'}")
    print("=" * 65)

    print("📥 Đang đọc dữ liệu từ Google Sheet (sheet 'Đang OFF')...")
    rows = load_sheet_data()
    print(f"✅ Đã tải {len(rows)} dòng dữ liệu.")

    am_data = parse_off_data(rows)
    print(f"📊 Tìm thấy {len(am_data)} AM có tuyến OFF.")

    if not am_data:
        print("⚠️ Không tìm thấy dữ liệu OFF tuyến hợp lệ!")
        return

    success_count = 0
    fail_count = 0

    print("\n📡 Đang xử lý tin nhắn cho từng AM...")
    for idx, (am_name, am_bcs) in enumerate(am_data.items(), 1):
        group_id = AM_GROUP_MAP.get(am_name)
        if not group_id:
            print(f"⚠️ AM '{am_name}' không tìm thấy ID Group ở danh sách cấu hình! (Bỏ qua)")
            continue

        am_msg = format_am_text_message(am_name, am_bcs)

        print(f"\n[{idx:02d}/{len(am_data):02d}] 👤 AM: {am_name:<22} | Group ID: {group_id}")

        if is_dry_run:
            print("  👉 [DRY-RUN] Tin nhắn văn bản mẫu dự kiến gửi:")
            print("  " + "-" * 55)
            print("  " + am_msg.replace('\n', '\n  '))
            print("  " + "-" * 55)
            success_count += 1
        else:
            ok_am, err_am = send_gtalk_text_message(group_id, am_msg, DEFAULT_TOKEN)
            if ok_am:
                print("  ✅ Gửi tin nhắn text thành công!")
                success_count += 1
            else:
                print(f"  ❌ Lỗi gửi text: {err_am}")
                fail_count += 1
            time.sleep(0.4)

    print("\n" + "=" * 65)
    print("📊 TỔNG KẾT BÁO CÁO:")
    print(f"   - Chế độ:     {'DRY-RUN (Chỉ xem trước)' if is_dry_run else 'THỰC HÀNH GỬI THẬT'}")
    print(f"   - Token:      {DEFAULT_TOKEN[:20]}...")
    print(f"   - Số AM gửi:  {success_count}/{len(am_data)}")
    if not is_dry_run:
        print(f"   - Thất bại:   {fail_count}/{len(am_data)}")
    else:
        print("💡 Để gửi thật, chạy lệnh: python push_off_tuyen_am.py --send")
    print("=" * 65)


if __name__ == "__main__":
    main()
