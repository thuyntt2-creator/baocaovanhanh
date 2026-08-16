import docx
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

doc_path = r"C:\Users\lap4all\Downloads\AOP_Hang_Nang_NTB_T7-T12_2026_v8_3.docx"

if not os.path.exists(doc_path):
    print("File không tồn tại")
    sys.exit(1)

doc = docx.Document(doc_path)
table = doc.tables[6]  # index 6 (Bảng 7)

print("=== ĐỌC CHI TIẾT BẢNG 7 ===")
for r_idx, row in enumerate(table.rows):
    row_text = [cell.text.strip().replace('\n', ' ') for cell in row.cells]
    print(f"Row {r_idx:02d}: {row_text}")
