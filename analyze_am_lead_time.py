# -*- coding: utf-8 -*-
"""
Script: analyze_am_lead_time.py
Tracks the average lead time (Thời gian gán đơn đầu tiên trung bình) for each AM.
Lead time is measured from the moment an order becomes aging (Entry BC time + 5 days)
until the AM triggers the first delivery attempt (num_deliver >= 1 or completed).
"""

import os
import sys
import json
import gspread
from datetime import datetime, timedelta
import pytz
import unicodedata
from google.oauth2.service_account import Credentials
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

try:
    sys.stdout.reconfigure(encoding='utf-8')
except AttributeError:
    pass

# ============ CONFIG & CONSTANTS ============
SHEET_KEY = '1WCzgao34cA_SttyB9ytHfE1qKTNl_3iFqDbEfw3lbyU'
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
TRACKER_FILE = os.path.join(BASE_DIR, 'am_lead_time_tracker.json')

scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

def parse_date(date_str):
    if not date_str:
        return None
    # Strip timezone info if any (+00:00)
    clean_str = date_str.split('+')[0].strip()
    for fmt in ('%Y-%m-%d %H:%M:%S.%f', '%Y-%m-%d %H:%M:%S', '%Y-%m-%d'):
        try:
            return datetime.strptime(clean_str, fmt)
        except ValueError:
            continue
    return None

def main():
    tz = pytz.timezone("Asia/Ho_Chi_Minh")
    now_dt = datetime.now(tz)
    now_naive = now_dt.replace(tzinfo=None)

    # 1. Connect to Google Sheets
    if not os.path.exists(JSON_FILE):
        print(f"❌ Không tìm thấy credentials.json")
        sys.exit(1)
        
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    sh = gc_client.open_by_key(SHEET_KEY)

    print("📖 Đọc dữ liệu từ tab 'data LM'...")
    ws_lm = sh.worksheet("data LM")
    lm_data = ws_lm.get_all_values()

    print("📖 Đọc dữ liệu từ tab 'Đơn giao aging trên 5 ngày'...")
    ws_aging = sh.worksheet("Đơn giao aging trên 5 ngày")
    aging_data = ws_aging.get_all_values()

    print("📖 Đọc dữ liệu từ tab 'Cơ cấu'...")
    ws_cocau = sh.worksheet("Cơ cấu")
    cocau_data = ws_cocau.get_all_values()

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

    # Extract backlog status from data LM
    lm_header = lm_data[1]
    lm_order_col = lm_header.index("Mã đơn hàng")
    lm_status_col = lm_header.index("Trạng thái")
    
    lm_status = {}
    for row in lm_data[2:]:
        if len(row) > max(lm_order_col, lm_status_col):
            m_don = row[lm_order_col].strip()
            t_thai = row[lm_status_col].strip()
            if m_don:
                lm_status[m_don] = t_thai

    # Headers indices
    aging_header = aging_data[0]
    ag_order_col = aging_header.index("order_code") if "order_code" in aging_header else aging_header.index("mã đơn")
    ag_bc_col = aging_header.index("bc")
    ag_id_bc_col = aging_header.index("id_bc")
    ag_time_col = aging_header.index("thoi_gian_nhap_bc")
    ag_am_col = aging_header.index("am_name")
    ag_num_col = aging_header.index("num_deliver") if "num_deliver" in aging_header else -1

    # Load tracker state
    tracker = {"active_tracking": {}, "history": []}
    if os.path.exists(TRACKER_FILE):
        try:
            with open(TRACKER_FILE, "r", encoding="utf-8") as f:
                tracker = json.load(f)
        except Exception as e:
            print(f"⚠️ Cảnh báo đọc tracker file: {e}")

    success_keywords = ['đã giao/ chuyển trả thành công', 'đã giao/chuyển trả thành công', 'n/a', 'thành công']
    current_active_unattempted = set()

    # Process all rows to check current status
    for row in aging_data[1:]:
        if len(row) > max(ag_order_col, ag_bc_col, ag_id_bc_col, ag_time_col, ag_am_col):
            order_code = row[ag_order_col].strip()
            if not order_code:
                continue

            status = lm_status.get(order_code, '#N/A')
            is_processed = status in ['#N/A', 'n/a'] or any(sk in status.lower() for sk in success_keywords)
            
            # Parse num_deliver
            num_deliver_val = row[ag_num_col].strip() if (ag_num_col != -1 and len(row) > ag_num_col) else '0'
            try:
                num_attempts = int(float(num_deliver_val))
            except ValueError:
                num_attempts = 0

            # Match AM name
            raw_am = row[ag_am_col].strip()
            if not raw_am or raw_am == '#N/A' or raw_am == '':
                bc_name = row[ag_bc_col].strip()
                id_bc = row[ag_id_bc_col].strip()
                am_name = cocau_map.get(id_bc, cocau_map.get(bc_name, "Không xác định"))
            else:
                am_name = raw_am
            am_name = unicodedata.normalize('NFC', am_name)

            # Aging start time (thoi_gian_nhap_bc + 5 days)
            thoi_gian_nhap_bc = row[ag_time_col].strip()
            entry_time = parse_date(thoi_gian_nhap_bc)
            if not entry_time:
                continue
            aging_start_time = entry_time + timedelta(days=5)

            # Case 1: Order is active and unattempted (num_attempts == 0)
            if not is_processed and num_attempts == 0:
                current_active_unattempted.add(order_code)
                # If not tracked yet, add to active tracking
                if order_code not in tracker["active_tracking"]:
                    tracker["active_tracking"][order_code] = {
                        "aging_start_time": aging_start_time.strftime('%Y-%m-%d %H:%M:%S'),
                        "am": am_name
                    }
            
            # Case 2: Order is in active tracking, but has now been attempted (num_attempts >= 1)
            elif order_code in tracker["active_tracking"] and num_attempts >= 1:
                # Transition detected! Calculate lead time
                start_str = tracker["active_tracking"][order_code]["aging_start_time"]
                start_dt = parse_date(start_str)
                if start_dt:
                    lead_time_seconds = (now_naive - start_dt).total_seconds()
                    lead_time_hours = max(0.0, lead_time_seconds / 3600.0)
                    
                    tracker["history"].append({
                        "code": order_code,
                        "am": am_name,
                        "aging_start_time": start_str,
                        "attempted_time": now_naive.strftime('%Y-%m-%d %H:%M:%S'),
                        "lead_time_hours": lead_time_hours
                    })
                # Remove from active tracking
                del tracker["active_tracking"][order_code]

    # Case 3: Orders in active tracking that are no longer in the active list at all
    # (meaning they got completed/delivered successfully without num_deliver update)
    inactive_tracked = [code for code in tracker["active_tracking"] if code not in current_active_unattempted]
    for order_code in inactive_tracked:
        # Check current status in data LM to see if completed
        status = lm_status.get(order_code, '#N/A')
        is_completed = status in success_keywords or status in ['#N/A', 'n/a']
        if is_completed:
            # Transition detected (completed)! Calculate lead time
            info = tracker["active_tracking"][order_code]
            start_str = info["aging_start_time"]
            start_dt = parse_date(start_str)
            if start_dt:
                lead_time_seconds = (now_naive - start_dt).total_seconds()
                lead_time_hours = max(0.0, lead_time_seconds / 3600.0)
                
                tracker["history"].append({
                    "code": order_code,
                    "am": info["am"],
                    "aging_start_time": start_str,
                    "attempted_time": now_naive.strftime('%Y-%m-%d %H:%M:%S'),
                    "lead_time_hours": lead_time_hours
                })
            # Remove from active tracking
            del tracker["active_tracking"][order_code]

    # Save state
    with open(TRACKER_FILE, "w", encoding="utf-8") as f:
        json.dump(tracker, f, ensure_ascii=False, indent=2)

    # 3. Calculate metrics and print leaderboard
    print("\n📊 BẢNG XẾP HẠNG THỜI GIAN GÁN ĐƠN ĐẦU TIÊN TRUNG BÌNH CỦA AM")
    print("=========================================================")
    
    # Group history by AM
    am_lead_times = {}
    for item in tracker["history"]:
        am = item["am"]
        if am not in am_lead_times:
            am_lead_times[am] = []
        am_lead_times[am].append(item["lead_time_hours"])

    if not am_lead_times:
        print("💡 Chưa có đơn hàng nào hoàn tất quá trình chuyển đổi để tính thời gian trung bình.")
        print("   (Hệ thống đã ghi nhận danh sách đơn đang theo dõi và sẽ tính toán ở các lượt chạy tiếp theo).")
        print(f"   Đang theo dõi: {len(tracker['active_tracking'])} đơn chưa giao của các AM.")
    else:
        leaderboard = []
        for am, times in am_lead_times.items():
            avg_hours = sum(times) / len(times)
            avg_days = avg_hours / 24.0
            leaderboard.append({
                "am": am,
                "avg_hours": avg_hours,
                "avg_days": avg_days,
                "count": len(times)
            })
            
        # Sort by avg_hours ascending (lower lead time is better!)
        leaderboard = sorted(leaderboard, key=lambda x: x["avg_hours"])
        
        for idx, item in enumerate(leaderboard):
            rank = idx + 1
            print(f"Hạng {rank}. AM {item['am']}: Trung bình gán sau {item['avg_days']:.1f} ngày ({item['avg_hours']:.1f} giờ) | Đã xử lý: {item['count']} đơn")

if __name__ == "__main__":
    main()
