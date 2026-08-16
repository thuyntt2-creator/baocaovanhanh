import os
import io
import sys
import gspread
from google.oauth2.service_account import Credentials

# Fix encoding cho Windows
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

def inspect_dang_off():
    print(f"📖 Connecting to sheet: {SHEET_KEY}...")
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    sh = gc_client.open_by_key(SHEET_KEY)
    
    ws = sh.worksheet("Đang OFF")
    
    # 1. Đọc công thức trong 10 dòng đầu
    print("\n📖 Đọc công thức trong 10 dòng đầu của tab 'Đang OFF'...")
    formulas = ws.get("A1:K10", value_render_option="FORMULA")
    for r_idx, row in enumerate(formulas):
        print(f"Dòng {r_idx + 1}: {row}")
        
    # 2. Đọc toàn bộ dữ liệu để kiểm tra cấu trúc và trùng lặp
    all_values = ws.get_all_values()
    print(f"\n📊 Tổng số dòng trong tab 'Đang OFF': {len(all_values)}")
    if len(all_values) < 2:
        return
        
    headers = [h.strip() for h in all_values[0]]
    print("Headers:", headers)
    
    # Kiểm tra trùng lặp ID phường/xã trong tab Đang OFF
    ward_id_col_idx = 6  # Mặc định cột G (index 6)
    for idx, h in enumerate(headers):
        if "id phường/xã" in h.lower() or "id phuong/xa" in h.lower():
            ward_id_col_idx = idx
            break
            
    print(f"Cột ID Phường/Xã được xác định là cột thứ {ward_id_col_idx + 1}: '{headers[ward_id_col_idx]}'")
    
    # Tìm các ID trùng lặp
    ward_counts = {}
    for r_idx, row in enumerate(all_values[1:]):
        row_num = r_idx + 2
        if len(row) <= ward_id_col_idx:
            continue
        w_id = row[ward_id_col_idx].strip()
        if not w_id:
            continue
            
        ward_name = row[5].strip() if len(row) > 5 else "Không rõ"
        district = row[4].strip() if len(row) > 4 else "Không rõ"
        province = row[3].strip() if len(row) > 3 else "Không rõ"
        
        info = {
            "row_num": row_num,
            "id": w_id,
            "ward": ward_name,
            "district": district,
            "province": province
        }
        if w_id not in ward_counts:
            ward_counts[w_id] = []
        ward_counts[w_id].append(info)
        
    duplicates = {k: v for k, v in ward_counts.items() if len(v) > 1}
    if duplicates:
        print(f"\n⚠️ Phát hiện {len(duplicates)} ID phường/xã bị trùng lặp trong nội bộ tab 'Đang OFF':")
        for w_id, infos in duplicates.items():
            print(f"  • ID {w_id}:")
            for info in infos:
                print(f"    - Dòng {info['row_num']}: {info['province']} | {info['district']} | {info['ward']}")
    else:
        print("\n✅ Không có ID phường/xã nào bị trùng lặp trong nội bộ tab 'Đang OFF'.")

if __name__ == "__main__":
    inspect_dang_off()
