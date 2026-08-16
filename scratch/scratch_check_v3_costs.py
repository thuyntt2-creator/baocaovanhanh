import openpyxl
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

file_path = r"C:\Users\lap4all\Downloads\V3_AOP_Hang_NTB_T7-T12_2026 mới.xlsx"
if not os.path.exists(file_path):
    print(f"File not found: {file_path}")
    sys.exit(1)

wb = openpyxl.load_workbook(file_path, data_only=False)
sheet = wb['Nguồn lực & chi phí']

print("=== Formulas/Values in V3 'Nguồn lực & chi phí' Sheet ===")
for r in range(1, 25):
    row_vals = []
    for c in range(1, 10):
        cell = sheet.cell(r, c)
        val = cell.value
        if isinstance(val, str) and val.startswith('='):
            row_vals.append(f"F:{val}")
        else:
            row_vals.append(val)
    if any(v is not None for v in row_vals):
        print(f"Row {r:2d}: {row_vals}")

wb.close()
