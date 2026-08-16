import os
import io
import sys
import re
import requests
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime

# Fix encoding cho Windows/Console/Task Scheduler
os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except AttributeError:
    pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
SHEET_KEY = '1PjzFqJO-wkQ8SNsPHD721_CbPr6c_ArZKuGGU6KqDZg'
scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

TELEGRAM_TOKEN = "8570130113:AAGXRiUaKBknVpgtm1_i9ZA47JRjAXmB21M"
TELEGRAM_CHAT_ID = "-5058464865"

def normalize_name(s):
    if not s:
        return ""
    s = s.strip().lower()
    s = re.sub(r'[\s\-\–\—\s]+', ' ', s)
    return s

def extract_date_prefix(s):
    if not s:
        return ""
    parts = s.strip().split()
    return parts[0] if parts else ""

def parse_dmy_date(s):
    s = s.strip()
    if not s:
        return None
    for fmt in ['%d-%m-%Y', '%d/%m/%Y']:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None

def parse_ymd_date(s):
    s = extract_date_prefix(s)
    if not s:
        return None
    try:
        return datetime.strptime(s, '%Y-%m-%d').date()
    except ValueError:
        return None

def escape_html(text):
    if not text:
        return ""
    return str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

def main():
    print(f"📖 Kết nối tới Google Sheet: {SHEET_KEY}...")
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    sh = gc_client.open_by_key(SHEET_KEY)
    
    # 1. Đọc danh sách bưu cục bất ổn
    print("🔒 Đọc danh sách bưu cục 'Bất ổn'...")
    ws_baton = sh.worksheet("Bất ổn")
    baton_data = ws_baton.get_all_values()
    
    unstable_hubs = set() # Set of normalized names
    baton_vung = {} # bc_norm -> vung
    for row in baton_data[1:]:
        # Cột U (index 20) là cột 'Trạng thái'
        if len(row) > 20 and row[20].strip() == "Bất ổn":
            if row[4].strip():
                unstable_hubs.add(normalize_name(row[4].strip()))
        if len(row) > 4 and row[4].strip():
            baton_vung[normalize_name(row[4].strip())] = row[1].strip()
            
    # 2. Đọc tab 'cơ cấu' để lấy danh sách tuyến thuộc bưu cục bất ổn
    print("🗺️ Tra cứu danh sách các tuyến thuộc bưu cục bất ổn...")
    ws_cocau = None
    for sname in ["CoCauVung", "cơ cấu", "Cơ cấu"]:
        try:
            ws_cocau = sh.worksheet(sname)
            break
        except Exception:
            pass
    if not ws_cocau:
        print("❌ Không tìm thấy tab CoCauVung hoặc cơ cấu trong spreadsheet.")
        return
    cocau_data = ws_cocau.get_all_values()
    
    warning_routes = [] # list of dicts
    ward_to_bc_map = {} # (t, h, w) -> bc_display
    for row in cocau_data[1:]:
        if len(row) < 8:
            continue
        ward_code = row[3].strip()
        tinh = row[0].strip()
        huyen = row[2].strip()
        phuong = row[4].strip()
        bc_new = row[7].strip()
        bc_old = row[6].strip()
        
        bc_display = bc_new if bc_new else bc_old
        bc_norm = normalize_name(bc_display)
        
        t_norm = normalize_name(tinh)
        h_norm = normalize_name(huyen)
        p_norm = normalize_name(phuong)
        ward_to_bc_map[(t_norm, h_norm, p_norm)] = bc_display
        
        if bc_norm in unstable_hubs and ward_code:
            warning_routes.append({
                "ward_code": ward_code,
                "tinh": tinh,
                "huyen": huyen,
                "phuong": phuong,
                "bc_display": bc_display
            })
            
    print(f"   • Tìm thấy {len(warning_routes)} tuyến thuộc {len(unstable_hubs)} bưu cục bất ổn.")

    # 3. Đọc dữ liệu từ tab 'vol'
    print("📦 Đọc dữ liệu sản lượng từ tab 'vol'...")
    ws_vol = sh.worksheet("vol")
    vol_data = ws_vol.get_all_values()
    
    # Tìm ngày mới nhất trong vol
    all_dates = []
    for row in vol_data[1:]:
        if len(row) > 0 and row[0].strip():
            d = extract_date_prefix(row[0])
            if d:
                all_dates.append(d)
    unique_dates = sorted(list(set(all_dates)), reverse=True)
    latest_date_str = unique_dates[0] if unique_dates else ""
    latest_date = parse_ymd_date(latest_date_str)
    print(f"   • Ngày dữ liệu mới nhất trong vol: {latest_date_str}")
    
    # Gom sản lượng Shopee của ngày mới nhất
    latest_vol_map = {} # (t, h, w) -> vol
    for row in vol_data[1:]:
        if len(row) < 8:
            continue
        d_str = extract_date_prefix(row[0])
        if d_str != latest_date_str:
            continue
        khach = row[6].strip()
        if khach not in ["Shopee-nhỏ", "Shopee-Bulky"]:
            continue
            
        t = normalize_name(row[2])
        h = normalize_name(row[3])
        w = normalize_name(row[4])
        
        try:
            v_val = float(row[7].strip().replace(",", ""))
        except ValueError:
            v_val = 0.0
            
        key = (t, h, w)
        latest_vol_map[key] = latest_vol_map.get(key, 0.0) + v_val

    # 4. Đọc dữ liệu 'Đang OFF'
    print("🔒 Đọc danh sách tuyến 'Đang OFF'...")
    ws_off = sh.worksheet("Đang OFF")
    off_data = ws_off.get_all_values()
    
    disabled_wards = {} # ward_code -> (start_date, end_date, row_data)
    for row in off_data[1:]:
        if len(row) < 10:
            continue
        tinh = row[0].strip()
        huyen = row[1].strip()
        phuong = row[2].strip()
        ward_code = row[3].strip()
        start_str = row[8].strip()
        end_str = row[9].strip()
        
        start_date = parse_dmy_date(start_str)
        end_date = parse_dmy_date(end_str)
        
        if ward_code and start_date and end_date:
            disabled_wards[ward_code] = {
                "start_date": start_date,
                "end_date": end_date,
                "tinh": tinh,
                "huyen": huyen,
                "phuong": phuong,
                "start_str": start_str,
                "end_str": end_str
            }

    # 5. Phân tích các cảnh báo
    today = datetime.now().date()
    
    chua_off_list = []
    het_han_off_list = []
    
    # Cảnh báo 1 & 2: Dành riêng cho tuyến thuộc bưu cục bất ổn
    for r in warning_routes:
        w_code = r["ward_code"]
        if w_code not in disabled_wards:
            chua_off_list.append(r)
        else:
            end_date = disabled_wards[w_code]["end_date"]
            if end_date <= today:
                het_han_off_list.append({
                    **r,
                    "end_date_str": end_date.strftime('%d-%m-%Y')
                })
                
    # Cảnh báo 3: Tuyến đang OFF nhưng có đơn phát sinh vào ngày mới nhất (Áp dụng cho tất cả các tuyến đang OFF)
    off_with_vol_list = []
    if latest_date:
        for w_code, off_info in disabled_wards.items():
            start_date = off_info["start_date"]
            end_date = off_info["end_date"]
            
            # Nếu ngày mới nhất nằm trong khoảng OFF
            if start_date <= latest_date <= end_date:
                t_norm = normalize_name(off_info["tinh"])
                h_norm = normalize_name(off_info["huyen"])
                p_norm = normalize_name(off_info["phuong"])
                
                vol = latest_vol_map.get((t_norm, h_norm, p_norm), 0.0)
                rounded_vol = int(round(vol))
                if rounded_vol > 0:
                    # Tra cứu bưu cục
                    bc_display = ward_to_bc_map.get((t_norm, h_norm, p_norm), "NTB")
                    off_with_vol_list.append({
                        "tinh": off_info["tinh"],
                        "huyen": off_info["huyen"],
                        "phuong": off_info["phuong"],
                        "ward_code": w_code,
                        "bc_display": bc_display,
                        "volume": rounded_vol,
                        "date_str": latest_date_str
                    })

    # 6. Gửi cảnh báo lên Telegram nếu phát hiện lỗi
    if not chua_off_list and not het_han_off_list and not off_with_vol_list:
        print("🎉 Tuyệt vời! Không phát hiện vấn đề nào cần cảnh báo.")
        return
        
    print(f"🚨 Phát hiện {len(chua_off_list)} tuyến chưa OFF, {len(het_han_off_list)} tuyến hết hạn OFF, và {len(off_with_vol_list)} tuyến OFF phát sinh đơn.")
    
    # Xây dựng nội dung tin nhắn dạng HTML để gửi Telegram
    today_str = datetime.now().strftime('%d-%m-%Y')
    msg_lines = [
        f"<b>⚠️ CẢNH BÁO TRẠNG THÁI OFF TUYẾN BẤT ỔN ({today_str}) ⚠️</b>\n"
    ]
    
    # Phần 1: Tuyến thuộc BC bất ổn nhưng chưa được OFF
    if chua_off_list:
        msg_lines.append("<b>🚨 1. TUYẾN THUỘC BC BẤT ỔN CHƯA ĐƯỢC OFF (Cần off gấp):</b>")
        by_bc = {}
        for r in chua_off_list:
            by_bc.setdefault(r["bc_display"], []).append(r)
            
        added_count = 0
        for bc, items in sorted(by_bc.items()):
            ward_desc_list = [f"{item['phuong']} ({item['ward_code']})" for item in items]
            ward_desc_str = ", ".join(ward_desc_list)
            
            line = f"  • <u>{escape_html(bc)}</u>: {len(items)} tuyến chưa OFF ({escape_html(ward_desc_str)})"
            
            current_len = sum(len(l) for l in msg_lines) + len(line) + 150
            if current_len > 3200:
                msg_lines.append(f"  • ... và <i>{len(by_bc) - added_count} bưu cục khác</i> chưa được OFF.")
                break
                
            msg_lines.append(line)
            added_count += 1
        msg_lines.append("")
        
    # Phần 2: Tuyến thuộc BC bất ổn hết hạn OFF
    if het_han_off_list:
        msg_lines.append("<b>⏰ 2. TUYẾN HẾT HẠN OFF NHƯNG BC VẪN BẤT ỔN (Cần gia hạn/gửi RQ mới):</b>")
        by_bc = {}
        for r in het_han_off_list:
            by_bc.setdefault(r["bc_display"], []).append(r)
            
        added_count = 0
        for bc, items in sorted(by_bc.items()):
            ward_desc_list = [f"{item['phuong']} (Hết hạn: {item['end_date_str']})" for item in items]
            ward_desc_str = ", ".join(ward_desc_list)
            
            line = f"  • <u>{escape_html(bc)}</u>: {len(items)} tuyến hết hạn OFF ({escape_html(ward_desc_str)})"
            
            current_len = sum(len(l) for l in msg_lines) + len(line) + 150
            if current_len > 3500:
                msg_lines.append(f"  • ... và <i>{len(by_bc) - added_count} bưu cục khác</i> đã hết hạn OFF.")
                break
                
            msg_lines.append(line)
            added_count += 1
        msg_lines.append("")

    # Phần 3: Tuyến đang OFF nhưng phát sinh đơn Shopee mới nhất (Mới bổ sung)
    if off_with_vol_list:
        msg_lines.append("<b>🔥 3. TUYẾN ĐANG OFF NHƯNG PHÁT SINH ĐƠN SHOPEE (Check ngay):</b>")
        by_bc = {}
        for r in off_with_vol_list:
            by_bc.setdefault(r["bc_display"], []).append(r)
            
        added_count = 0
        for bc, items in sorted(by_bc.items()):
            ward_desc_list = [f"{item['phuong']} (<b>{item['volume']} đơn</b>)" for item in items]
            ward_desc_str = ", ".join(ward_desc_list)
            
            line = f"  • <u>{escape_html(bc)}</u>: {len(items)} tuyến có đơn ({escape_html(ward_desc_str)})"
            
            current_len = sum(len(l) for l in msg_lines) + len(line) + 150
            if current_len > 3900:
                msg_lines.append(f"  • ... và <i>{len(by_bc) - added_count} bưu cục khác</i> phát sinh đơn lỗi.")
                break
                
            msg_lines.append(line)
            added_count += 1
        msg_lines.append("")
        
    msg_lines.append("👉 <i>Vui lòng kiểm tra và xử lý gấp các tuyến phát sinh đơn khi đang OFF!</i>")
    
    full_message = "\n".join(msg_lines)
    
    # Gửi qua API Telegram
    telegram_url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    payload = {
        "chat_id": TELEGRAM_CHAT_ID,
        "text": full_message,
        "parse_mode": "HTML"
    }
    
    print("Sending message to Telegram...")
    response = requests.post(telegram_url, json=payload)
    if response.status_code == 200:
        print("🎉 Gửi tin nhắn cảnh báo Telegram thành công!")
    else:
        print(f"❌ Gửi tin nhắn thất bại. Mã phản hồi: {response.status_code}")
        print(response.text)

if __name__ == "__main__":
    main()
