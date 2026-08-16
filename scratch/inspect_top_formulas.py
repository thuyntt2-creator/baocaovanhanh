import openpyxl
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

xlsx_aop_path = r"C:\Users\lap4all\Downloads\AOP_Hang_NTB_T7-T12_2026.xlsx"
wb = openpyxl.load_workbook(xlsx_aop_path, data_only=False)

for sname in ['Volume Giao', 'Volume Lấy']:
    sheet = wb[sname]
    print(f"\n=== Formulas in {sname} (Rows 3-18) ===")
    for r_idx in range(3, 19):
        row_vals = []
        for c_idx in range(1, 10):
            cell = sheet.cell(r_idx, c_idx)
            val = cell.value
            if isinstance(val, str) and val.startswith('='):
                row_vals.append(f"F:{val}")
            else:
                row_vals.append(val)
        print(f"Row {r_idx:2d}: {row_vals}")

