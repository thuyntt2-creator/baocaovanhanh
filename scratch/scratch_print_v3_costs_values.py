import openpyxl
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

file_path = r"C:\Users\lap4all\Downloads\V3_AOP_Hang_NTB_T7-T12_2026 mới.xlsx"
wb = openpyxl.load_workbook(file_path, data_only=True)
sheet = wb['Nguồn lực & chi phí']

print("=== Evaluated Values in V3 'Nguồn lực & chi phí' Sheet ===")
for r in range(1, 23):
    row_vals = [sheet.cell(r, c).value for c in range(1, 8)]
    if any(v is not None for v in row_vals):
        print(f"Row {r:2d} | {str(row_vals[0]):<45} | {row_vals[1:]}")

wb.close()
