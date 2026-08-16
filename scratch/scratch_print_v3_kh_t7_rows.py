import openpyxl
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

file_path = r"C:\Users\lap4all\Downloads\V3_AOP_Hang_NTB_T7-T12_2026 mới.xlsx"
wb = openpyxl.load_workbook(file_path, data_only=True)
sheet = wb['Kế hoạch T7']

print("=== Rows 11-50 in Kế hoạch T7 ===")
for r in range(11, 51):
    row_vals = [sheet.cell(r, c).value for c in range(1, 10)]
    if any(v is not None for v in row_vals):
        print(f"Row {r:3d}: {row_vals}")

wb.close()
