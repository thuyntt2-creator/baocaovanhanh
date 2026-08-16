import os
import io
import sys
import json
import re
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
HISTORY_FILE = os.path.join(BASE_DIR, 'baton_warning_history.json')
scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

def normalize_name(s):
    if not s:
        return ""
    # Chuẩn hóa khoảng trắng, chữ thường, thay thế các loại gạch ngang đặc biệt thành gạch ngang thường
    s = s.strip().lower()
    s = re.sub(r'[\s\-\–\—\s]+', ' ', s)
    return s

def extract_date_prefix(s):
    if not s:
        return ""
    parts = s.strip().split()
    return parts[0] if parts else ""

def load_history():
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"⚠️ Lỗi đọc file lịch sử JSON, khởi tạo mới: {e}")
    return {}

def main():
    print(f"📖 Kết nối tới Google Sheet: {SHEET_KEY}...")
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    sh = gc_client.open_by_key(SHEET_KEY)
    
    # 1. Đọc dữ liệu từ tab 'plan rq'
    print("⚡ Đọc dữ liệu từ tab 'plan rq'...")
    ws_plan = sh.worksheet("plan rq")
    plan_data = ws_plan.get_all_values()
    
    # Đọc tab Bất ổn để lấy map Vùng cho bưu cục
    print("⚡ Đọc tab 'Bất ổn' để lấy bản đồ vùng của các bưu cục...")
    ws_baton = sh.worksheet("Bất ổn")
    baton_data = ws_baton.get_all_values()
    baton_vung_map = {}
    for row in baton_data[1:]:
        if len(row) > 3 and row[3].strip():
            baton_vung_map[row[3].strip()] = row[1].strip() if len(row) > 1 else "NTB"
            
    # Lấy danh sách bưu cục cần OFF từ plan rq
    target_hubs = {} # ID Bưu cục: {name, vung}
    for row in plan_data[1:]:
        if len(row) < 2:
            continue
        hub_id = row[0].strip()
        hub_name = row[1].strip()
        if not hub_id:
            continue
        vung = baton_vung_map.get(hub_id, "NTB")
        target_hubs[hub_id] = {
            "name": hub_name,
            "vung": vung
        }
        
    print(f"📋 Tìm thấy {len(target_hubs)} bưu cục cần OFF được chỉ định trong tab 'plan rq'.")
    if not target_hubs:
        print("ℹ️ Tab 'plan rq' trống. Không có bưu cục nào được chỉ định để xử lý.")
        return
        
    # Load lịch sử cảnh báo (chỉ để đọc số ngày dồn cảnh báo)
    history = load_history()
    
    # Xác định tháng N và tháng N-1 để đếm ngày cảnh báo
    run_dt = datetime.now()
    snapshot_date_str = run_dt.strftime('%Y-%m-%d')
    year_n = run_dt.year
    month_n = run_dt.month
    
    # Tính tháng N-1
    if month_n == 1:
        month_n1 = 12
        year_n1 = year_n - 1
    else:
        month_n1 = month_n - 1
        year_n1 = year_n
        
    print(f"📊 Tính toán số ngày dồn cảnh báo cho Tháng N ({month_n:02d}/{year_n}) và Tháng N-1 ({month_n1:02d}/{year_n1})...")

    # 2. Đọc danh sách đang OFF để loại trừ
    print("🔒 Đọc danh sách tuyến 'Đang OFF' để loại trừ...")
    ws_off = sh.worksheet("Đang OFF")
    off_data = ws_off.get_all_values()
    off_ward_id_idx = 3 # ID phường/xã (Col 4)
    disabled_wards = set()
    for row in off_data[1:]:
        if len(row) > off_ward_id_idx and row[off_ward_id_idx].strip():
            disabled_wards.add(row[off_ward_id_idx].strip())
    print(f"ℹ️ Có {len(disabled_wards)} tuyến đã tắt hoàn toàn.")

    # 3. Đọc tab 'cơ cấu' để lấy danh sách phường xã của bưu cục trong plan rq
    print("🗺️ Đọc tab 'cơ cấu'...")
    ws_cocau = sh.worksheet("cơ cấu")
    cocau_data = ws_cocau.get_all_values()
    # Headers: ['Tỉnh', 'ID Huyện', 'Huyện/TP', 'ward_code', 'Phường/Xã', 'ID BC', 'Bưu cục', 'Bưu Cục new', 'AM']
    
    proposed_routes = [] # list of dicts
    for row in cocau_data[1:]:
        if len(row) < 8:
            continue
        ward_code = row[3].strip()
        if not ward_code or ward_code in disabled_wards:
            continue
            
        tinh = row[0].strip()
        huyen = row[2].strip()
        phuong_xa = row[4].strip()
        id_bc = row[5].strip()
        buu_cuc_goc = row[6].strip()
        buu_cuc_new = row[7].strip()
        
        # So khớp bưu cục cơ cấu với bưu cục trong plan rq
        matched_hub = None
        if buu_cuc_new and buu_cuc_new in [h["name"] for h in target_hubs.values()]:
            # Tìm ID bưu cục tương ứng
            for hid, hinfo in target_hubs.items():
                if hinfo["name"] == buu_cuc_new:
                    matched_hub = (hid, hinfo)
                    break
        elif id_bc and id_bc in target_hubs:
            matched_hub = (id_bc, target_hubs[id_bc])
            
        if matched_hub:
            hid, hinfo = matched_hub
            
            # Đếm ngày cảnh báo từ lịch sử JSON
            warning_dates = history.get(hid, {}).get("warning_dates", [])
            days_n = 0
            days_n1 = 0
            for d_str in warning_dates:
                try:
                    dt = datetime.strptime(d_str, '%Y-%m-%d')
                    if dt.year == year_n and dt.month == month_n:
                        days_n += 1
                    elif dt.year == year_n1 and dt.month == month_n1:
                        days_n1 += 1
                except Exception:
                    pass
            
            proposed_routes.append({
                "vung": hinfo["vung"],
                "tinh": tinh,
                "huyen": huyen,
                "phuong_xa": phuong_xa,
                "ward_code": ward_code,
                "id_bc": id_bc,
                "buu_cuc_name": buu_cuc_new if buu_cuc_new else buu_cuc_goc,
                "buu_cuc_new": buu_cuc_new,
                "days_n": days_n,
                "days_n1": days_n1
            })
            
    print(f"📋 Tìm thấy {len(proposed_routes)} tuyến thuộc các bưu cục cần OFF được chỉ định.")
    if not proposed_routes:
        print("ℹ️ Không tìm thấy tuyến chưa OFF nào thuộc các bưu cục đã nhập.")
        return

    # 4. Tính sản lượng Shopee trung bình 7 ngày từ tab 'vol'
    print("📦 Đọc tab 'vol' và tính toán sản lượng Shopee 7 ngày...")
    ws_vol = sh.worksheet("vol")
    vol_data = ws_vol.get_all_values()
    # Headers: ['Date', 'Vùng', 'Tỉnh', 'Quận/Huyện', 'Phường/Xã', 'Bưu cục', 'Khách hàng', 'Volume']
    
    # Tìm 7 ngày cuối cùng có trong dữ liệu
    all_dates = []
    for row in vol_data[1:]:
        if len(row) > 0 and row[0].strip():
            d = extract_date_prefix(row[0])
            if d:
                all_dates.append(d)
    unique_dates_vol = sorted(list(set(all_dates)))
    last_7_dates = unique_dates_vol[-7:] if len(unique_dates_vol) >= 7 else unique_dates_vol
    print(f"   • Sử dụng dữ liệu volume của 7 ngày cuối: {last_7_dates}")
    
    vol_map = {} # (tinh, huyen, ward) -> total_shopee_vol
    for row in vol_data[1:]:
        if len(row) < 8:
            continue
        d_val = extract_date_prefix(row[0])
        if not d_val or d_val not in last_7_dates:
            continue
        khach = row[6].strip()
        if khach not in ["Shopee-nhỏ", "Shopee-Bulky"]:
            continue
            
        t = normalize_name(row[2])
        h = normalize_name(row[3])
        w = normalize_name(row[4])
        
        try:
            v_val = float(row[7].strip().replace(",", ""))
        except Exception:
            v_val = 0.0
            
        key = (t, h, w)
        vol_map[key] = vol_map.get(key, 0.0) + v_val
        
    # Gán sản lượng vào proposed_routes
    for route in proposed_routes:
        key = (normalize_name(route["tinh"]), normalize_name(route["huyen"]), normalize_name(route["phuong_xa"]))
        total_vol = vol_map.get(key, 0.0)
        route["shopee_avg_7day"] = int(round(total_vol / 7.0)) if len(last_7_dates) > 0 else 0

    # 5. Tính trung bình % GTC 3 ngày gần nhất (không tính T7-CN) từ tab 'PFM'
    print("📈 Đọc tab 'PFM' và tính GTC bưu cục 3 ngày gần nhất...")
    ws_pfm = sh.worksheet("PFM")
    pfm_data = ws_pfm.get_all_values()
    # Headers: ['Cấp Quản Lý', 'Chi tiết', 'Loại Hàng', 'Time', 'Volume', '% Gán', '% GTC', ...]
    
    # Tìm 3 ngày làm việc gần nhất (loại bỏ T7, CN)
    pfm_dates = []
    for row in pfm_data[1:]:
        if len(row) > 3 and row[3].strip():
            d_val = extract_date_prefix(row[3])
            if d_val:
                pfm_dates.append(d_val)
    unique_pfm_dates = sorted(list(set(pfm_dates)))
    
    working_dates = []
    for d_val in unique_pfm_dates:
        if not d_val:
            continue
        try:
            dt = datetime.strptime(d_val, '%Y-%m-%d')
            # 5 là Thứ 7, 6 là Chủ nhật
            if dt.weekday() not in [5, 6]:
                working_dates.append(d_val)
        except Exception:
            pass
            
    last_3_working_days = working_dates[-3:] if len(working_dates) >= 3 else working_dates
    print(f"   • Sử dụng dữ liệu PFM của 3 ngày làm việc gần nhất: {last_3_working_days}")
    
    # Nhóm % GTC theo bưu cục trong 3 ngày này
    gtc_map = {} # pfm_bc_name -> list of floats
    for row in pfm_data[1:]:
        if len(row) < 7 or not row[3].strip():
            continue
        t_val = extract_date_prefix(row[3])
        if not t_val or t_val not in last_3_working_days:
            continue
            
        bc_detail = normalize_name(row[1])
        gtc_str = row[6].strip().replace("%", "")
        try:
            gtc_val = float(gtc_str)
        except Exception:
            continue
            
        if bc_detail not in gtc_map:
            gtc_map[bc_detail] = []
        gtc_map[bc_detail].append(gtc_val)
        
    # Gán % GTC trung bình vào proposed_routes
    for route in proposed_routes:
        bc_key_new = normalize_name(route["buu_cuc_new"])
        bc_key_goc = normalize_name(route["buu_cuc_name"])
        
        gtc_vals = None
        if bc_key_new in gtc_map:
            gtc_vals = gtc_map[bc_key_new]
        elif bc_key_goc in gtc_map:
            gtc_vals = gtc_map[bc_key_goc]
            
        if gtc_vals:
            avg_gtc = sum(gtc_vals) / len(gtc_vals)
            route["gtc_avg_3day"] = f"{round(avg_gtc, 2)}%"
        else:
            route["gtc_avg_3day"] = "0%"

    # 6. Đọc bảo toàn dữ liệu thủ công của tab 'Note RQ'
    print("📝 Kiểm tra dữ liệu cũ trong 'Note RQ' để bảo toàn nhập liệu thủ công...")
    ws_note = sh.worksheet("Note RQ")
    note_data = ws_note.get_all_values()
    
    # Lưu trữ các dòng cũ theo ward_code
    existing_records = {}
    
    if len(note_data) > 1:
        for row in note_data[1:]:
            if len(row) > 5 and row[5].strip():
                w_code = row[5].strip()
                has_manual = False
                manual_indices = [9, 10, 11, 15, 16, 17] # J, K, L, P, Q, R
                for idx in manual_indices:
                    if idx < len(row) and row[idx].strip():
                        has_manual = True
                        break
                
                if has_manual:
                    existing_records[w_code] = row

    # 7. Tổng hợp các dòng dữ liệu để ghi
    final_rows = []
    written_wards = set()
    input_date_dmy = run_dt.strftime('%d-%m-%Y')
    
    # Thêm các dòng đề xuất mới từ plan rq
    for route in proposed_routes:
        w_code = route["ward_code"]
        written_wards.add(w_code)
        
        old_row = existing_records.get(w_code, [])
        new_row = [""] * 18
        
        # Điền các cột tự động tính toán
        new_row[0] = input_date_dmy                # Ngày input
        new_row[1] = route["vung"]                  # Vùng
        new_row[2] = route["tinh"]                  # Tỉnh
        new_row[3] = route["huyen"]                 # Quận/ huyện
        new_row[4] = route["phuong_xa"]             # Phường/xã cần tắt
        new_row[5] = w_code                         # ID phường/xã
        new_row[6] = route["id_bc"]                 # ID Bưu cục
        new_row[7] = route["buu_cuc_name"]          # Bưu Cục
        new_row[8] = str(route["shopee_avg_7day"])  # SL Shopee TB 7 ngày
        
        # Bảo toàn cột thủ công 10, 11, 12
        if len(old_row) > 9: new_row[9] = old_row[9]
        if len(old_row) > 10: new_row[10] = old_row[10]
        if len(old_row) > 11: new_row[11] = old_row[11]
        
        new_row[12] = route["gtc_avg_3day"]         # GTC 3 ngày gần nhất
        new_row[13] = str(route["days_n1"])         # Số ngày nằm trong cảnh báo tháng N-1
        new_row[14] = str(route["days_n"])          # Số ngày nằm trong cảnh báo tháng N
        
        # Bảo toàn cột thủ công 16, 17, 18
        if len(old_row) > 15: new_row[15] = old_row[15]
        if len(old_row) > 16: new_row[16] = old_row[16]
        if len(old_row) > 17: new_row[17] = old_row[17]
        
        final_rows.append(new_row)
        
    # Giữ lại các dòng cũ có dữ liệu thủ công
    for w_code, old_row in existing_records.items():
        if w_code not in written_wards:
            padded_row = old_row + [""] * (18 - len(old_row))
            final_rows.append(padded_row[:18])
            
    print(f"📊 Tổng số dòng chuẩn bị ghi vào 'Note RQ' (bao gồm cả dòng cũ đã bảo toàn): {len(final_rows)}")
    
    # Ghi dữ liệu vào Google Sheet
    try:
        print("✍️ Đang thử ghi toàn bộ 18 cột vào tab 'Note RQ'...")
        ws_note.batch_clear(["A2:R1000"])
        if final_rows:
            ws_note.update(f"A2:R{1 + len(final_rows)}", [r[:18] for r in final_rows])
        print("🎉 Đã cập nhật thành công toàn bộ 18 cột vào tab 'Note RQ'!")
    except gspread.exceptions.APIError as e:
        if "protected cell or object" in str(e):
            print("⚠️ Cảnh báo: Phát hiện cột hoặc ô bị khóa bảo vệ (cột B, C, D, E chứa Vùng, Tỉnh, Huyện, Phường).")
            print("🔄 Thực hiện chế độ Fallback: Chỉ ghi các cột không bị khóa (A và F-R)...")
            
            col_a_vals = [[r[0]] for r in final_rows]
            col_f_r_vals = [r[5:18] for r in final_rows]
            
            ws_note.batch_clear(["A2:A1000", "F2:S1000"])
            if col_a_vals:
                ws_note.update(f"A2:A{1 + len(final_rows)}", col_a_vals)
            if col_f_r_vals:
                ws_note.update(f"F2:R{1 + len(final_rows)}", col_f_r_vals)
                
            print("🎉 Đã cập nhật thành công dữ liệu vào tab 'Note RQ' (chế độ Fallback, bỏ qua các cột bị khóa).")
        else:
            raise e

if __name__ == "__main__":
    main()
