import docx
import shutil
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

# Đường dẫn file
docx_path = r"C:\Users\lap4all\Downloads\AOP_Hang_Nang_NTB_T7-T12_2026_v8_1.docx"
output_path = r"C:\Users\lap4all\Downloads\AOP_Hang_Nang_NTB_T7-T12_2026_v8_2.docx"

if not os.path.exists(docx_path):
    print(f"File không tồn tại: {docx_path}")
    sys.exit(1)

doc = docx.Document(docx_path)

# Dữ liệu chuẩn từ Excel V18
# Đơn vị: triệu VNĐ (giá trị gốc chia cho 1e6)
months = ['T7', 'T8', 'T9', 'T10', 'T11', 'T12']

# 1. Dữ liệu bảng chi phí chung (Bảng 9)
cost_data = {
    'T7':  {'xe': 781.2,  'giao': 345.0, 'kho_ql': 120.0, 'rent': 82.5,  'tong': 1328.7},
    'T8':  {'xe': 855.6,  'giao': 375.0, 'kho_ql': 120.0, 'rent': 87.6,  'tong': 1438.2},
    'T9':  {'xe': 1116.0, 'giao': 495.0, 'kho_ql': 120.0, 'rent': 125.1, 'tong': 1856.1},
    'T10': {'xe': 1190.4, 'giao': 510.0, 'kho_ql': 120.0, 'rent': 127.5, 'tong': 1947.9},
    'T11': {'xe': 1332.0, 'giao': 585.0, 'kho_ql': 120.0, 'rent': 152.55, 'tong': 2189.55},
    'T12': {'xe': 1413.6, 'giao': 600.0, 'kho_ql': 120.0, 'rent': 155.1, 'tong': 2288.7},
}

# 2. Dữ liệu bảng chi tiết bưu cục (Bảng 10)
hub_data = {
    'BCCK Nha Trang': {'T7': 499.8, 'T8': 552.0, 'T9': 882.0, 'T10': 917.4, 'T11': 1014.0, 'T12': 1074.0},
    'BCCK Đơn Dương':  {'T7': 201.6, 'T8': 238.8, 'T9': 249.0, 'T10': 253.8, 'T11': 285.0, 'T12': 306.0},
    'BCCK Di Linh':    {'T7': 343.2, 'T8': 358.2, 'T9': 402.0, 'T10': 410.4, 'T11': 489.0, 'T12': 499.8},
    'BCCK Đức Linh':   {'T7': 201.6, 'T8': 201.6, 'T9': 234.0, 'T10': 238.8, 'T11': 249.0, 'T12': 253.8},
    'TỔNG':            {'T7': 1246.2, 'T8': 1350.6, 'T9': 1767.0, 'T10': 1820.4, 'T11': 2037.0, 'T12': 2133.6}
}

# Tiến hành cập nhật các bảng
for idx, table in enumerate(doc.tables):
    first_col_text = [row.cells[0].text.strip() for row in table.rows]
    
    # --- Cập nhật Bảng 9 (Bảng chi phí chung) ---
    if any('chi phí xe tải' in h.lower() for h in first_col_text):
        print(f"Cập nhật Bảng {idx} (Chi phí chung)...")
        for row in table.rows:
            if len(row.cells) < 7:
                continue
            header = row.cells[0].text.strip().lower()
            if 'chi phí xe tải' in header:
                for col_idx, m in enumerate(months):
                    row.cells[col_idx+1].text = f"{cost_data[m]['xe']:.1f}"
            elif 'chi phí nv giao hàng' in header:
                for col_idx, m in enumerate(months):
                    row.cells[col_idx+1].text = f"{cost_data[m]['giao']:.1f}"
            elif 'chi phí nv kho & quản lý' in header or 'chi phí nv kho' in header:
                row.cells[0].text = "Chi phí NV kho & quản lý (8 người cố định, triệu đ)"
                for col_idx, m in enumerate(months):
                    row.cells[col_idx+1].text = f"{cost_data[m]['kho_ql']:.1f}"
            elif 'chi phí thuê mặt bằng' in header:
                for col_idx, m in enumerate(months):
                    row.cells[col_idx+1].text = f"{cost_data[m]['rent']:.2f}".rstrip('0').rstrip('.')
            elif 'tổng chi phí' in header:
                for col_idx, m in enumerate(months):
                    row.cells[col_idx+1].text = f"{cost_data[m]['tong']:.2f}".rstrip('0').rstrip('.')

    # --- Cập nhật Bảng 10 (Chi tiết từng bưu cục) ---
    # Phải đảm bảo đây là bảng chi phí chi tiết bằng cách kiểm tra tiêu đề dòng đầu có 'triệu đ' hoặc 'tháng'
    is_cost_detail_table = False
    if len(table.rows) > 0:
        row0_text = [cell.text.lower() for cell in table.rows[0].cells]
        if any('triệu đ/tháng' in t or 'triệu đ' in t for t in row0_text):
            is_cost_detail_table = True
            
    if is_cost_detail_table:
        print(f"Cập nhật Bảng {idx} (Chi tiết bưu cục)...")
        for row in table.rows:
            if len(row.cells) < 7:
                continue
            header = row.cells[0].text.strip()
            
            # Map tên bưu cục
            matched_key = None
            for key in hub_data.keys():
                if key.lower() in header.lower() or (key == 'TỔNG' and 'tổng' in header.lower()):
                    matched_key = key
                    break
                    
            if matched_key:
                print(f"  Ghi đè dòng: '{header}' -> key: '{matched_key}'")
                for col_idx, m in enumerate(months):
                    row.cells[col_idx+1].text = f"{hub_data[matched_key][m]:.1f}"

# Lưu lại file đè trực tiếp lên file gốc để người dùng kiểm tra
doc.save(output_path)
print(f"Cập nhật hoàn tất! File đã được lưu ra: {output_path}")
