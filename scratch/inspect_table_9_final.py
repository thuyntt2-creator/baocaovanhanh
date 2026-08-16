import docx
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

doc_path = r"C:\Users\lap4all\Downloads\AOP_Hang_Nang_NTB_T7-T12_2026_v8_3.docx"

if not os.path.exists(doc_path):
    print("File không tồn tại")
    sys.exit(1)

doc = docx.Document(doc_path)
table = doc.tables[8]

print("=== IN ROW 5 THỰC TẾ ===")
row_text = [cell.text.strip() for cell in table.rows[5].cells]
for idx, text in enumerate(row_text):
    print(f"Cell {idx}: {text}")
