import docx
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

doc_path = r"C:\Users\lap4all\Downloads\AOP_Hang_Nang_NTB_T7-T12_2026_v8_1.docx"

if not os.path.exists(doc_path):
    print("File không tồn tại")
    sys.exit(1)

doc = docx.Document(doc_path)
table = doc.tables[2]  # index 2 (Bảng 3)

print("=== ĐỌC CHI TIẾT BẢNG 3 ===")
for r_idx, row in enumerate(table.rows):
    cells_text = [cell.text.strip().replace('\n', ' ') for cell in row.cells]
    # Rút gọn merged cells
    cleaned = []
    for t in cells_text:
        if not cleaned or cleaned[-1] != t:
            cleaned.append(t)
    print(f"Row {r_idx:02d}: {cleaned}")
