import openpyxl
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

xlsx_aop_path = r"C:\Users\lap4all\Downloads\AOP_Hang_NTB_T7-T12_2026.xlsx"
wb = openpyxl.load_workbook(xlsx_aop_path, data_only=False)
sheet = wb['Mật độ & phương án']

print("=== Formulas in Mật độ & phương án ===")
for r in range(1, 20):
    row_vals = []
    for c_idx in range(1, 10):
        cell = sheet.cell(r, c_idx)
        val = cell.value
        if isinstance(val, str) and val.startswith('='):
            row_vals.append(f"F:{val}")
        else:
            row_vals.append(val)
    if any(v is not None for v in row_vals):
        print(f"Row {r:2d}: {row_vals}")

