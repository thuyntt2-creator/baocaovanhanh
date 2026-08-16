import sys
import io
import os
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from datetime import datetime

# Fix encoding for Windows Command Prompt / Tasks Scheduler
os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except AttributeError:
    pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
SHEET_KEY = '1sUboaLTIeNTsbG56Re70-Xt5M7pHd8r4Neh-k4vzFt4'

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def get_day_of_week_vn(dt):
    days_vn = {
        0: 'Thứ 2',
        1: 'Thứ 3',
        2: 'Thứ 4',
        3: 'Thứ 5',
        4: 'Thứ 6',
        5: 'Thứ 7',
        6: 'Chủ Nhật'
    }
    return days_vn.get(dt.weekday(), '')

def main():
    print("🔄 Bắt đầu tiến trình ánh xạ và cập nhật dữ liệu...")
    
    if not os.path.exists(JSON_FILE):
        print(f"❌ Không tìm thấy file credentials tại: {JSON_FILE}")
        sys.exit(1)
        
    try:
        # 1. Kết nối Google Sheets
        print("🔑 Đang kết nối tới Google Sheets API...")
        creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
        gc_client = gspread.authorize(creds)
        sh = gc_client.open_by_key(SHEET_KEY)
        print(f"✅ Đã kết nối thành công tới bảng tính: '{sh.title}'")
        
        # 2. Tải dữ liệu từ sheet "Cơ cấu"
        print("📖 Đọc dữ liệu từ sheet 'Cơ cấu'...")
        ws_cocau = sh.worksheet("Cơ cấu")
        cocau_data = ws_cocau.get_all_values()
        if len(cocau_data) < 2:
            print("❌ Lỗi: Sheet 'Cơ cấu' không có dữ liệu.")
            sys.exit(1)
            
        cocau_header = cocau_data[0]
        try:
            cc_code_idx = cocau_header.index("Mã bưu cục")
            cc_name_idx = cocau_header.index("Bưu cục")
            cc_prov_idx = cocau_header.index("Tỉnh")
        except ValueError:
            cc_code_idx = 0
            cc_name_idx = 1
            cc_prov_idx = 3
            
        cc_map = {}
        for row in cocau_data[1:]:
            if len(row) > max(cc_code_idx, cc_name_idx, cc_prov_idx):
                code = row[cc_code_idx].strip()
                name = row[cc_name_idx].strip()
                prov = row[cc_prov_idx].strip()
                if code:
                    cc_map[code] = {
                        'bc': name,
                        'prov': prov
                    }
        print(f"   Đã đọc {len(cc_map)} bưu cục từ Cơ cấu.")
        
        # 3. Tải dữ liệu từ sheet "data thô"
        print("📖 Đọc dữ liệu từ sheet 'data thô'...")
        ws_raw = sh.worksheet("data thô")
        raw_data = ws_raw.get_all_values()
        if len(raw_data) < 2:
            print("❌ Lỗi: Sheet 'data thô' trống.")
            sys.exit(1)
            
        raw_header = raw_data[0]
        try:
            raw_cols = {
                'WarehouseID': raw_header.index('WarehouseID'),
                'Ngay': raw_header.index('Ngay'),
                'LoaiHang': raw_header.index('LoaiHang'),
                'Ca': raw_header.index('Ca'),
                'Volume': raw_header.index('Volume'),
                'OntimeAdd': raw_header.index('OntimeAdd'),
                'OntimeDeliver': raw_header.index('OntimeDeliver'),
                'OntimeReturn': raw_header.index('OntimeReturn'),
                'Leadtime': raw_header.index('Leadtime'),
                'VolumeLeadtime': raw_header.index('VolumeLeadtime'),
            }
        except ValueError as e:
            print(f"❌ Lỗi: Cột dữ liệu thiếu trong sheet 'data thô'. Chi tiết: {e}")
            sys.exit(1)
            
        print(f"   Đang xử lý {len(raw_data) - 1} dòng dữ liệu thô...")
        
        aggregated_data = {}
        skipped_count = 0
        
        for idx, row in enumerate(raw_data[1:], start=2):
            if len(row) <= max(raw_cols.values()):
                continue
                
            w_id = row[raw_cols['WarehouseID']].strip()
            
            # Kiểm tra và ánh xạ trực tiếp từ sheet Cơ cấu
            if w_id not in cc_map:
                skipped_count += 1
                continue
                
            chi_tiet = cc_map[w_id]['bc']
            prov = cc_map[w_id]['prov']
            cap_quan_ly = f"NTB - {prov}" if prov else "NTB"
                
            # Lấy Loại Hàng
            loai_hang = f"{row[raw_cols['LoaiHang']].strip()} {row[raw_cols['Ca']].strip()}"
            
            # Lấy và format Time
            ngay_str = row[raw_cols['Ngay']].strip()
            try:
                dt = datetime.strptime(ngay_str, "%Y-%m-%d")
                thu_str = get_day_of_week_vn(dt)
                time_str = f"{ngay_str} - {thu_str}"
            except ValueError:
                time_str = ngay_str
                
            # Cast các trường số
            try:
                vol = float(row[raw_cols['Volume']]) if row[raw_cols['Volume']] else 0.0
                ontime_add = float(row[raw_cols['OntimeAdd']]) if row[raw_cols['OntimeAdd']] else 0.0
                ontime_deliver = float(row[raw_cols['OntimeDeliver']]) if row[raw_cols['OntimeDeliver']] else 0.0
                ontime_return = float(row[raw_cols['OntimeReturn']]) if row[raw_cols['OntimeReturn']] else 0.0
                lt = float(row[raw_cols['Leadtime']]) if row[raw_cols['Leadtime']] else 0.0
                vol_lt = float(row[raw_cols['VolumeLeadtime']]) if row[raw_cols['VolumeLeadtime']] else 0.0
            except ValueError:
                # Bỏ qua dòng bị lỗi kiểu số
                continue
                
            key = (cap_quan_ly, chi_tiet, loai_hang, time_str)
            if key not in aggregated_data:
                aggregated_data[key] = {
                    'volume': 0.0,
                    'ontime_add': 0.0,
                    'ontime_deliver': 0.0,
                    'ontime_return': 0.0,
                    'leadtime_sum': 0.0,
                    'volume_leadtime_sum': 0.0
                }
                
            aggregated_data[key]['volume'] += vol
            aggregated_data[key]['ontime_add'] += ontime_add
            aggregated_data[key]['ontime_deliver'] += ontime_deliver
            aggregated_data[key]['ontime_return'] += ontime_return
            aggregated_data[key]['leadtime_sum'] += lt
            aggregated_data[key]['volume_leadtime_sum'] += vol_lt
            
        print(f"   Đã bỏ qua {skipped_count} dòng (do ngoài vùng NTB hoặc không map được).")
        print(f"   Số lượng nhóm sau gộp: {len(aggregated_data)} nhóm.")
        
        # 4. Xây dựng dữ liệu ghi xuống sheet DB
        output_rows = []
        for key, vals in aggregated_data.items():
            cap_quan_ly, chi_tiet, loai_hang, time_str = key
            volume = vals['volume']
            
            pct_gan = vals['ontime_add'] / volume if volume > 0 else 0.0
            pct_gtc = vals['ontime_deliver'] / volume if volume > 0 else 0.0
            pct_return = vals['ontime_return'] / volume if volume > 0 else 0.0
            
            avg_leadtime = vals['leadtime_sum'] / vals['volume_leadtime_sum'] if vals['volume_leadtime_sum'] > 0 else 0.0
            
            output_rows.append([
                cap_quan_ly,
                chi_tiet,
                loai_hang,
                time_str,
                volume,
                pct_gan,
                pct_gtc,
                pct_return,
                avg_leadtime
            ])
            
        # Sắp xếp lại theo Ngày tăng dần, sau đó là Cấp Quản Lý, Bưu cục, Loại hàng
        def sort_key(r):
            date_part = r[3].split(" - ")[0]
            return (date_part, r[0], r[1], r[2])
            
        output_rows.sort(key=sort_key)
        
        # 5. Ghi đè vào sheet "DB"
        print("📊 Đang ghi đè dữ liệu vào sheet 'DB'...")
        ws_data = sh.worksheet("DB")
        
        # Xóa các dữ liệu cũ từ A2:I để tránh dư thừa (chỉ xóa cột A đến I từ dòng 2 trở đi)
        print("🧹 Đang dọn dẹp các ô dữ liệu cũ (A2:I)...")
        ws_data.batch_clear(["A2:I100000"])
        
        # Ghi đè dữ liệu mới
        print(f"📤 Đang cập nhật {len(output_rows)} dòng dữ liệu vào các cột A đến I...")
        ws_data.update(
            range_name=f"A2:I{len(output_rows) + 1}",
            values=output_rows,
            value_input_option='USER_ENTERED'
        )
        print("✔️ Cập nhật thành công dữ liệu thô vào sheet 'DB'!")
        
    except Exception as e:
        print(f"❌ Xảy ra lỗi bất ngờ: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
