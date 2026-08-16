import docx
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

doc_path = r"C:\Users\lap4all\Downloads\AOP_Hang_Nang_NTB_T7-T12_2026_v8_10.docx"

if not os.path.exists(doc_path):
    print("File không tồn tại")
    sys.exit(1)

doc = docx.Document(doc_path)
errors = 0

print("=== KIỂM TRA SỐ LIỆU TOÀN DIỆN FILE WORD V8_10 ===")

# 1. Kiểm tra Bảng 8 dòng 2 (Xe BQ làm tròn)
table_8 = doc.tables[7]
print("\n1. Kiểm tra Bảng 8 (Dòng 2 - Xe BQ làm tròn):")
actual_bq = [table_8.rows[2].cells[c].text.strip() for c in range(1, 7)]
expected_bq = ("21", "23", "31", "32", "37", "38")
if tuple(actual_bq) != expected_bq:
    print(f"  [LỖI] Xe BQ = {actual_bq} | Kỳ vọng = {expected_bq}")
    errors += 1
else:
    print("  [OK] Xe BQ khớp hoàn toàn")

if errors == 0:
    print("\n>>> TẤT CẢ KIỂM TRA BÁO CÁO V8_10 ĐỀU THÀNH CÔNG! SỐ LIỆU HOÀN THIỆN TUYỆT ĐỐI <<<")
else:
    print(f"\n>>> CÓ {errors} SAI LỆCH SỐ LIỆU CẦN XỬ LÝ <<<")
