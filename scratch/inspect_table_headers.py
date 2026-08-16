import docx
import sys

sys.stdout.reconfigure(encoding='utf-8')

docx_path = r"C:\Users\lap4all\Downloads\AOP_Hang_Nang_NTB_T7-T12_2026_v8_1.docx"
doc = docx.Document(docx_path)

# Tìm bảng chi phí
# Nó có các hàng như 'Chi phí xe tải 1.9T'
target_table = None
for idx, table in enumerate(doc.tables):
    for row in table.rows:
        if any('chi phí xe tải' in cell.text.lower() for cell in row.cells):
            target_table = table
            print(f"Tìm thấy bảng chi phí ở vị trí Bảng {idx}")
            break
    if target_table:
        break

if target_table:
    for r_idx, row in enumerate(target_table.rows):
        cells_text = [cell.text.strip().replace('\n', ' ') for cell in row.cells]
        print(f"Row {r_idx}: {cells_text}")
else:
    print("Không tìm thấy bảng chi phí.")
