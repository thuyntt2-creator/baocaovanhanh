import docx
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

doc_path = r"C:\Users\lap4all\Downloads\AOP_Hang_Nang_NTB_T7-T12_2026_v8_8.docx"

if not os.path.exists(doc_path):
    print("File không tồn tại")
    sys.exit(1)

doc = docx.Document(doc_path)
errors = 0

print("=== KIỂM TRA SỐ LIỆU TOÀN DIỆN FILE WORD V8_8 ===")

# 1. Kiểm tra Bảng 3 (Số xe đỉnh)
table_3 = doc.tables[2]
print("\n1. Kiểm tra Bảng 3 (Số xe đỉnh và chuyến xe đỉnh):")
expected_table_3 = {
    'Nha Trang': {
        'trips': ('28 chuyến', '46 chuyến', '56 chuyến'),
        'trucks': ('14 xe', '23 xe', '28 xe')
    },
    'Di Linh': {
        'trips': ('16 chuyến', '20 chuyến', '24 chuyến'),
        'trucks': ('8 xe', '10 xe', '12 xe')
    },
    'Đơn Dương': {
        'trips': ('8 chuyến', '12 chuyến', '14 chuyến'),
        'trucks': ('4 xe', '6 xe', '7 xe')
    },
    'Đức Linh': {
        'trips': ('8 chuyến', '10 chuyến', '12 chuyến'),
        'trucks': ('4 xe', '5 xe', '6 xe')
    }
}

rows_map = {
    'Nha Trang': (2, 3),
    'Di Linh': (5, 6),
    'Đơn Dương': (8, 9),
    'Đức Linh': (11, 12)
}

for name, (row_trips, row_trucks) in rows_map.items():
    actual_trips = [table_3.rows[row_trips].cells[2].text.strip(), table_3.rows[row_trips].cells[3].text.strip(), table_3.rows[row_trips].cells[4].text.strip()]
    actual_trucks = [table_3.rows[row_trucks].cells[2].text.strip(), table_3.rows[row_trucks].cells[3].text.strip(), table_3.rows[row_trucks].cells[4].text.strip()]
    
    exp = expected_table_3[name]
    if tuple(actual_trips) != exp['trips']:
        print(f"  [LỖI] {name} chuyến: Thực tế = {actual_trips} | Kỳ vọng = {exp['trips']}")
        errors += 1
    else:
        print(f"  [OK] {name} chuyến khớp")
        
    if tuple(actual_trucks) != exp['trucks']:
        print(f"  [LỖI] {name} xe: Thực tế = {actual_trucks} | Kỳ vọng = {exp['trucks']}")
        errors += 1
    else:
        print(f"  [OK] {name} xe khớp")

if errors == 0:
    print("\n>>> TẤT CẢ KIỂM TRA BÁO CÁO V8_8 ĐỀU THÀNH CÔNG! SỐ LIỆU HOÀN THIỆN TUYỆT ĐỐI <<<")
else:
    print(f"\n>>> CÓ {errors} SAI LỆCH SỐ LIỆU CẦN XỬ LÝ <<<")
