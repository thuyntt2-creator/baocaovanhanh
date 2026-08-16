import docx
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

doc_path = r"C:\Users\lap4all\Downloads\AOP_Hang_Nang_NTB_T7-T12_2026_v8_1.docx"

if not os.path.exists(doc_path):
    print("File không tồn tại")
    sys.exit(1)

doc = docx.Document(doc_path)

print("=== TÌM BẢNG CHỨA GAP HOẶC CHUYẾN ===")
for t_idx, table in enumerate(doc.tables):
    for r_idx, row in enumerate(table.rows):
        for cell in row.cells:
            text = cell.text.lower()
            if 'gap' in text or 'chuyến' in text:
                print(f"Khớp tại BẢNG {t_idx+1}, hàng {r_idx+1}: {cell.text.strip().replace('\n', ' ')}")
                break
