import docx
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

doc_path = r"C:\Users\lap4all\Downloads\AOP_Hang_Nang_NTB_T7-T12_2026_v8_1.docx"

if not os.path.exists(doc_path):
    print("File không tồn tại")
    sys.exit(1)

doc = docx.Document(doc_path)

tables_to_inspect = [5, 6, 7, 9, 10]  # 0-based index tương ứng Bảng 6, 7, 8, 10, 11

for t_idx in tables_to_inspect:
    table = doc.tables[t_idx]
    print(f"\n=== BẢNG {t_idx+1} (Rows: {len(table.rows)}, Cols: {len(table.columns)}) ===")
    for r_idx, row in enumerate(table.rows):
        cells_text = [cell.text.strip().replace('\n', ' ') for cell in row.cells]
        print(f"  Row {r_idx:02d}: {cells_text}")
