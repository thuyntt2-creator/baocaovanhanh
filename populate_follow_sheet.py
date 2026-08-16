import os
import io
import sys
import re
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
from gspread_formatting import *

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

def main():
    print(f"📖 Kết nối tới Google Sheet: {SHEET_KEY}...")
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    sh = gc_client.open_by_key(SHEET_KEY)
    
    # 1. Đọc dữ liệu Đang OFF
    print("🔒 Đọc danh sách tuyến 'Đang OFF'...")
    ws_off = sh.worksheet("Đang OFF")
    off_data = ws_off.get_all_values()
    
    off_routes = []
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
        
        if tinh and huyen and phuong:
            off_routes.append({
                "tinh": tinh,
                "huyen": huyen,
                "phuong": phuong,
                "ward_code": ward_code,
                "start_date": start_date,
                "end_date": end_date,
                "start_str": start_str,
                "end_str": end_str
            })
    print(f"   • Tìm thấy {len(off_routes)} tuyến trong danh sách Đang OFF.")

    # 2. Đọc dữ liệu từ tab 'vol'
    print("📦 Đọc dữ liệu sản lượng từ tab 'vol'...")
    ws_vol = sh.worksheet("vol")
    vol_data = ws_vol.get_all_values()
    
    # Gom danh sách ngày hiện có trong dữ liệu volume
    all_dates = []
    for row in vol_data[1:]:
        if len(row) > 0 and row[0].strip():
            d = extract_date_prefix(row[0])
            if d:
                all_dates.append(d)
    unique_dates = sorted(list(set(all_dates)), reverse=True)
    print(f"   • Danh sách ngày trong vol (mới nhất đầu): {unique_dates}")

    # Gom sản lượng Shopee (Shopee-nhỏ & Shopee-Bulky) theo (tinh, huyen, phuong, date_str)
    vol_map = {} # (t, h, w, date_str) -> volume
    for row in vol_data[1:]:
        if len(row) < 8:
            continue
        d_str = extract_date_prefix(row[0])
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
            
        key = (t, h, w, d_str)
        vol_map[key] = vol_map.get(key, 0.0) + v_val

    # 3. Phân tích vùng từ tab Bất ổn & cơ cấu
    print("🗺️ Đọc thông tin vùng từ tab Bất ổn & cơ cấu...")
    ws_baton = sh.worksheet("Bất ổn")
    baton_data = ws_baton.get_all_values()
    baton_vung = {}
    for row in baton_data[1:]:
        if len(row) > 4 and row[3].strip():
            baton_vung[normalize_name(row[4].strip())] = row[1].strip()
            
    ws_cocau = sh.worksheet("cơ cấu")
    cocau_data = ws_cocau.get_all_values()
    ward_to_bc_map = {}
    for row in cocau_data[1:]:
        if len(row) >= 8:
            t = normalize_name(row[0])
            h = normalize_name(row[2])
            w = normalize_name(row[4])
            ward_to_bc_map[(t, h, w)] = normalize_name(row[7].strip())

    # 4. Duyệt các tuyến đang OFF và lọc ra những tuyến bị phát sinh đơn khi OFF
    print("🔍 Đối chiếu dữ liệu phát sinh đơn hàng trong thời gian OFF...")
    follow_records = []
    for route in off_routes:
        t_norm = normalize_name(route["tinh"])
        h_norm = normalize_name(route["huyen"])
        p_norm = normalize_name(route["phuong"])
        
        start_date = route["start_date"]
        end_date = route["end_date"]
        
        # Xác định vùng cho tuyến
        bc_name = ward_to_bc_map.get((t_norm, h_norm, p_norm), "")
        vung = baton_vung.get(bc_name, "NTB")
        
        has_off_volume = False
        date_volumes = {}
        
        for d_str in unique_dates:
            d_date = parse_ymd_date(d_str)
            # Kiểm tra xem ngày có nằm trong khoảng OFF của tuyến không
            is_off_day = False
            if start_date and end_date and d_date:
                is_off_day = (start_date <= d_date <= end_date)
                
            vol = vol_map.get((t_norm, h_norm, p_norm, d_str), 0.0)
            
            if is_off_day:
                # Nằm trong thời gian OFF, lấy volume (làm tròn số nguyên)
                rounded_vol = int(round(vol))
                date_volumes[d_str] = rounded_vol
                if rounded_vol > 0:
                    has_off_volume = True
            else:
                # Không nằm trong thời gian OFF, để trống
                date_volumes[d_str] = ""
                
        if has_off_volume:
            follow_records.append({
                "vung": vung,
                "tinh": route["tinh"],
                "huyen": route["huyen"],
                "phuong": route["phuong"],
                "start_str": route["start_str"],
                "end_str": route["end_str"],
                "volumes": date_volumes
            })
            
    print(f"   • Phát hiện {len(follow_records)} tuyến đang OFF nhưng vẫn phát sinh đơn Shopee.")

    # 5. Ghi dữ liệu vào tab Follow
    print("✍️ Ghi dữ liệu và định dạng tab 'Follow'...")
    ws_follow = sh.worksheet("Follow")
    
    # Clear dữ liệu cũ
    ws_follow.clear()
    
    # Xây dựng ma trận ghi
    # Dòng 1: Tiêu đề
    row1 = ["Theo dõi Volume khu vực OFF/Capdown", "", "", "", "", "", "OrderDate"]
    
    # Dòng 2: Headers
    headers = ["Vùng", "Tỉnh", "Quận/huyện", "Phường/xã", "Thời gian tắt", "Thời gian mở"] + unique_dates
    
    rows_to_write = [row1, headers]
    
    # Sắp xếp các tuyến theo Vùng -> Tỉnh -> Huyện -> Phường cho đẹp mắt
    follow_records = sorted(follow_records, key=lambda x: (x["vung"], x["tinh"], x["huyen"], x["phuong"]))
    
    for r in follow_records:
        row_data = [r["vung"], r["tinh"], r["huyen"], r["phuong"], r["start_str"], r["end_str"]]
        for d_str in unique_dates:
            row_data.append(r["volumes"].get(d_str, ""))
        rows_to_write.append(row_data)
        
    # Ghi toàn bộ dữ liệu xuống sheet
    ws_follow.update(f"A1:{gspread.utils.rowcol_to_a1(len(rows_to_write), len(headers))}", rows_to_write)
    
    # 6. Thiết lập đóng băng (Freeze) & Định dạng mỹ thuật
    print("🎨 Áp dụng định dạng mỹ thuật (Freeze, Fonts, Colors)...")
    ws_follow.freeze(rows=2, cols=6)
    
    # Định dạng tiêu đề cột G1 (OrderDate) & A1 (Tiêu đề bảng)
    title_fmt = CellFormat(
        textFormat=TextFormat(bold=True, italic=True, foregroundColor=Color(0.85, 0.1, 0.1), fontSize=12)
    )
    orderdate_fmt = CellFormat(
        textFormat=TextFormat(bold=True, italic=True, fontSize=10)
    )
    format_cell_range(ws_follow, 'A1:A1', title_fmt)
    format_cell_range(ws_follow, 'G1:G1', orderdate_fmt)
    
    # Định dạng Header Row 2
    header_fmt = CellFormat(
        backgroundColor=Color(0.92, 0.92, 0.92),
        textFormat=TextFormat(bold=True),
        horizontalAlignment='CENTER'
    )
    max_col_letter = gspread.utils.rowcol_to_a1(1, len(headers))[:-1]
    format_cell_range(ws_follow, f'A2:{max_col_letter}2', header_fmt)
    
    # Định dạng có điều kiện (Conditional Formatting) sử dụng Custom Formula để bỏ qua ô trống
    # Cột bắt đầu từ G (cột 7), dòng từ 3 đến hết
    num_rows = len(rows_to_write)
    data_range = GridRange.from_a1_range(f"G3:{max_col_letter}{num_rows}", ws_follow)
    
    # Luật Xanh: Volume >= 10
    rule_green = ConditionalFormatRule(
        ranges=[data_range],
        booleanRule=BooleanRule(
            condition=BooleanCondition(type='CUSTOM_FORMULA', values=[{'userEnteredValue': '=AND(ISNUMBER(G3), G3>=10)'}]),
            format=CellFormat(
                backgroundColor=Color(0.85, 0.94, 0.85),
                textFormat=TextFormat(foregroundColor=Color(0.1, 0.45, 0.1), bold=True)
            )
        )
    )
    
    # Luật Hồng: Volume < 10 và >= 0 (loại trừ các ô trống)
    rule_pink = ConditionalFormatRule(
        ranges=[data_range],
        booleanRule=BooleanRule(
            condition=BooleanCondition(type='CUSTOM_FORMULA', values=[{'userEnteredValue': '=AND(ISNUMBER(G3), G3>=0, G3<10)'}]),
            format=CellFormat(
                backgroundColor=Color(0.98, 0.86, 0.86),
                textFormat=TextFormat(foregroundColor=Color(0.75, 0.15, 0.15))
            )
        )
    )
    
    # Lưu luật định dạng
    rules = get_conditional_format_rules(ws_follow)
    rules.clear()
    rules.append(rule_green)
    rules.append(rule_pink)
    rules.save()
    
    print("🎉 Hoàn tất tự động hóa cập nhật tab 'Follow'!")

if __name__ == "__main__":
    main()
