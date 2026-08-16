import os
import io
import sys
import re
import gspread
from google.oauth2.service_account import Credentials
from gspread_formatting import *
from datetime import datetime
import calendar

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

def overlaps_month(start_date, end_date, year, month):
    if not start_date or not end_date:
        return False
    _, last_day = calendar.monthrange(year, month)
    m_start = datetime(year, month, 1).date()
    m_end = datetime(year, month, last_day).date()
    return max(start_date, m_start) <= min(end_date, m_end)

def main():
    print(f"📖 Kết nối tới Google Sheet: {SHEET_KEY}...")
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    sh = gc_client.open_by_key(SHEET_KEY)
    
    # 1. Đọc dữ liệu lịch sử từ tab 'raw '
    print("🔒 Đọc lịch sử OFF tuyến từ tab 'raw '...")
    ws_raw = sh.worksheet("raw ")
    raw_data = ws_raw.get_all_values()
    
    # 2. Đọc tab 'cơ cấu' để lấy mapping ward_code -> Bưu cục mới
    print("🗺️ Đọc cơ cấu tổ chức để chuẩn hóa thông tin bưu cục...")
    ws_cocau = sh.worksheet("cơ cấu")
    cocau_data = ws_cocau.get_all_values()
    
    # ward_code -> {tinh, huyen, phuong, id_bc, bc_old, bc_new}
    cocau_map = {}
    for row in cocau_data[1:]:
        if len(row) < 8:
            continue
        tinh = row[0].strip()
        huyen = row[2].strip()
        phuong = row[4].strip()
        ward_code = row[3].strip()
        id_bc = row[5].strip()
        bc_old = row[6].strip()
        bc_new = row[7].strip()
        
        if ward_code:
            cocau_map[ward_code] = {
                "tinh": tinh,
                "huyen": huyen,
                "phuong": phuong,
                "id_bc": id_bc,
                "bc_old": bc_old,
                "bc_new": bc_new
            }
            
    # 3. Đọc thông tin Vùng từ tab 'Bất ổn'
    print("🌍 Lấy thông tin vùng từ tab 'Bất ổn'...")
    ws_baton = sh.worksheet("Bất ổn")
    baton_data = ws_baton.get_all_values()
    baton_vung = {}
    for row in baton_data[1:]:
        if len(row) > 4 and row[3].strip():
            baton_vung[normalize_name(row[4].strip())] = row[1].strip()

    # 4. Gom nhóm và tính toán tần suất OFF theo ID phường/xã dựa trên khoảng thời gian tắt/mở
    print("📊 Đang phân tích và gom nhóm dữ liệu lịch sử...")
    ward_agg = {} # ward_code -> {month_counts: {month -> count}, raw_fallback_info, total_rows}
    
    for row in raw_data[1:]:
        if len(row) < 18:
            continue
        vung = row[2].strip()
        tinh = row[3].strip()
        huyen = row[4].strip()
        phuong = row[5].strip()
        ward_code = row[6].strip()
        id_bc = row[7].strip()
        bc_name = row[8].strip()
        
        start_str = row[16].strip()
        end_str = row[17].strip()
        start_date = parse_dmy_date(start_str)
        end_date = parse_dmy_date(end_str)
        
        if not ward_code:
            continue
            
        if ward_code not in ward_agg:
            ward_agg[ward_code] = {
                "vung": vung,
                "tinh": tinh,
                "huyen": huyen,
                "phuong": phuong,
                "id_bc": id_bc,
                "bc_name": bc_name,
                "month_counts": {4: 0, 5: 0, 6: 0, 7: 0},
                "total_rows": 0
            }
            
        info = ward_agg[ward_code]
        info["total_rows"] += 1
        
        # Kiểm tra xem khoảng ngày tắt có giao với các tháng 4, 5, 6, 7 hay không
        for m in [4, 5, 6, 7]:
            if overlaps_month(start_date, end_date, 2026, m):
                info["month_counts"][m] += 1

    # 5. Xây dựng ma trận dữ liệu ghi xuống tab historyoff
    print("✍️ Chuẩn bị ma trận dữ liệu và ghi xuống tab 'historyoff'...")
    ws_hist = sh.worksheet("historyoff")
    
    # Clear dữ liệu cũ
    ws_hist.clear()
    
    # Dòng 1: Tiêu đề
    row1 = ["Bảng tổng hợp tần suất OFF tuyến khu vực NTB", "", "", "", "", "", "", "", "", "", "", ""]
    # Dòng 2: Headers
    headers = [
        "Vùng", "Tỉnh", "Quận/huyện", "Phường/xã", "Mã phường/xã", 
        "Mã bưu cục", "Tên bưu cục", 
        "Tháng 4", "Tháng 5", "Tháng 6", "Tháng 7", "Tổng cộng"
    ]
    
    rows_to_write = [row1, headers]
    
    # Xây dựng các hàng dữ liệu
    data_rows = []
    for w_code, info in ward_agg.items():
        # Ưu tiên lấy thông tin chuẩn hóa từ cơ cấu
        cc_info = cocau_map.get(w_code)
        if cc_info:
            tinh = cc_info["tinh"]
            huyen = cc_info["huyen"]
            phuong = cc_info["phuong"]
            id_bc = cc_info["id_bc"]
            # Sử dụng Bưu cục new (tên ngắn gọn)
            bc_display = cc_info["bc_new"] if cc_info["bc_new"] else cc_info["bc_old"]
        else:
            tinh = info["tinh"]
            huyen = info["huyen"]
            phuong = info["phuong"]
            id_bc = info["id_bc"]
            bc_display = info["bc_name"]
            
        # Tìm vùng
        bc_norm = normalize_name(bc_display)
        vung = baton_vung.get(bc_norm, info["vung"] if info["vung"] else "NTB")
        
        # Lấy count của từng tháng
        counts = info["month_counts"]
        c4 = counts[4]
        c5 = counts[5]
        c6 = counts[6]
        c7 = counts[7]
        total = info["total_rows"]
        
        data_rows.append({
            "vung": vung,
            "tinh": tinh,
            "huyen": huyen,
            "phuong": phuong,
            "ward_code": w_code,
            "id_bc": id_bc,
            "bc_display": bc_display,
            "c4": c4,
            "c5": c5,
            "c6": c6,
            "c7": c7,
            "total": total
        })
        
    # Sắp xếp theo thứ tự Vùng -> Tỉnh -> Huyện -> Phường
    data_rows = sorted(data_rows, key=lambda x: (x["vung"], x["tinh"], x["huyen"], x["phuong"]))
    
    # Thêm vào rows_to_write
    for r in data_rows:
        rows_to_write.append([
            r["vung"], r["tinh"], r["huyen"], r["phuong"], r["ward_code"],
            r["id_bc"], r["bc_display"],
            r["c4"], r["c5"], r["c6"], r["c7"], r["total"]
        ])
        
    # Ghi dữ liệu xuống Google Sheet
    ws_hist.update(f"A1:{gspread.utils.rowcol_to_a1(len(rows_to_write), len(headers))}", rows_to_write)
    
    # 6. Thiết lập đóng băng & Định dạng mỹ thuật
    print("🎨 Áp dụng định dạng mỹ thuật (Freeze, Fonts, Colors)...")
    ws_hist.freeze(rows=2, cols=7) # Đóng băng 2 dòng tiêu đề, 7 cột thông tin
    
    # Định dạng tiêu đề chính cell A1
    title_fmt = CellFormat(
        textFormat=TextFormat(bold=True, italic=True, foregroundColor=Color(0.1, 0.3, 0.7), fontSize=12)
    )
    format_cell_range(ws_hist, 'A1:A1', title_fmt)
    
    # Định dạng Header Row 2
    header_fmt = CellFormat(
        backgroundColor=Color(0.91, 0.94, 0.98), # Premium Google Blue tint background
        textFormat=TextFormat(bold=True),
        horizontalAlignment='CENTER'
    )
    max_col_letter = gspread.utils.rowcol_to_a1(1, len(headers))[:-1]
    format_cell_range(ws_hist, f'A2:{max_col_letter}2', header_fmt)
    
    # Định dạng căn giữa cột số liệu (cột H đến L)
    num_rows = len(rows_to_write)
    middle_fmt = CellFormat(horizontalAlignment='CENTER')
    format_cell_range(ws_hist, f'H3:L{num_rows}', middle_fmt)
    
    # Định dạng chữ đậm cho cột L (Tổng cộng)
    total_col_fmt = CellFormat(textFormat=TextFormat(bold=True), horizontalAlignment='CENTER')
    format_cell_range(ws_hist, f'L3:L{num_rows}', total_col_fmt)
    
    # Định dạng có điều kiện (Conditional Formatting) cho cột L (Tổng cộng)
    data_range = GridRange.from_a1_range(f"L3:L{num_rows}", ws_hist)
    
    # Luật Xanh: Tổng số lần OFF từ 1 đến 2 lần (Ổn định tương đối)
    rule_green = ConditionalFormatRule(
        ranges=[data_range],
        booleanRule=BooleanRule(
            condition=BooleanCondition(type='CUSTOM_FORMULA', values=[{'userEnteredValue': '=AND(ISNUMBER(L3), L3>=1, L3<=2)'}]),
            format=CellFormat(
                backgroundColor=Color(0.9, 0.96, 0.91), # Xanh lá nhạt
                textFormat=TextFormat(foregroundColor=Color(0.08, 0.45, 0.2)) # Chữ xanh đậm
            )
        )
    )
    
    # Luật Đỏ: Tổng số lần OFF từ 3 lần trở lên (Rất bất ổn)
    rule_red = ConditionalFormatRule(
        ranges=[data_range],
        booleanRule=BooleanRule(
            condition=BooleanCondition(type='CUSTOM_FORMULA', values=[{'userEnteredValue': '=AND(ISNUMBER(L3), L3>=3)'}]),
            format=CellFormat(
                backgroundColor=Color(0.99, 0.91, 0.9), # Đỏ/hồng nhạt
                textFormat=TextFormat(foregroundColor=Color(0.77, 0.13, 0.12), bold=True) # Chữ đỏ đậm, in đậm
            )
        )
    )
    
    # Lưu luật định dạng vào trang tính
    rules = get_conditional_format_rules(ws_hist)
    rules.clear()
    rules.append(rule_green)
    rules.append(rule_red)
    rules.save()
    
    print("🎉 Hoàn tất tự động hóa cập nhật tab 'historyoff'!")

if __name__ == "__main__":
    main()
