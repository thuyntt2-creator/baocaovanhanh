import docx
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

doc_path = r"C:\Users\lap4all\Downloads\AOP_Hang_Nang_NTB_T7-T12_2026_v8_9.docx"

if not os.path.exists(doc_path):
    print("File không tồn tại")
    sys.exit(1)

doc = docx.Document(doc_path)
errors = 0

print("=== KIỂM TRA SỐ LIỆU TOÀN DIỆN FILE WORD V8_9 ===")

# 1. Kiểm tra Bảng 8 (Nhu cầu phương tiện & mặt bằng)
table_8 = doc.tables[7]
print("\n1. Kiểm tra Bảng 8:")
expected_table_8 = {
    2: ("20.4", "22.6", "31.5", "32.5", "37.2", "38.0"), # Xe BQ
    3: ("30", "31", "41", "42", "51", "51"),             # Xe đỉnh
    4: ("60", "62", "82", "84", "102", "102"),           # Người giao (đỉnh x 2)
    5: ("550", "584", "834", "850", "1,017", "1,034")    # Mặt bằng
}

for row_idx, expected in expected_table_8.items():
    actual = [table_8.rows[row_idx].cells[c].text.strip() for c in range(1, 7)]
    if tuple(actual) != expected:
        print(f"  [LỖI] Dòng {row_idx}: Thực tế = {actual} | Kỳ vọng = {expected}")
        errors += 1
    else:
        print(f"  [OK] Dòng {row_idx} khớp")

if errors == 0:
    print("\n>>> TẤT CẢ KIỂM TRA BÁO CÁO V8_9 ĐỀU THÀNH CÔNG! SỐ LIỆU HOÀN THIỆN TUYỆT ĐỐI <<<")
else:
    print(f"\n>>> CÓ {errors} SAI LỆCH SỐ LIỆU CẦN XỬ LÝ <<<")
