import docx
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

doc_path = r"C:\Users\lap4all\Downloads\AOP_Hang_Nang_NTB_T7-T12_2026_v8_3.docx"

if not os.path.exists(doc_path):
    print("File không tồn tại")
    sys.exit(1)

doc = docx.Document(doc_path)

print("=== TÌM BẢNG CHỨA TUYỂN DỤNG ===")
for t_idx, table in enumerate(doc.tables):
    for r_idx, row in enumerate(table.rows):
        cells_text = [cell.text.strip().replace('\n', ' ') for cell in row.cells]
        if any('tuyển' in c.lower() or 'nhân sự' in c.lower() for c in cells_text):
            print(f"Bảng {t_idx+1}, Hàng {r_idx:02d}: {cells_text}")
