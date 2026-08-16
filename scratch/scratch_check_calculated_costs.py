import openpyxl
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

file_path = r"C:\Users\lap4all\Downloads\AOP_Hang_NTB_T7-T12_2026_calculated.xlsx"
if not os.path.exists(file_path):
    print(f"File not found: {file_path}")
    sys.exit(1)

wb = openpyxl.load_workbook(file_path, data_only=True)
sheet = wb['Nguồn lực & chi phí']

print("=== Values in Calculated 'Nguồn lực & chi phí' ===")
for r in range(1, 20):
    row_vals = [sheet.cell(r, c).value for c in range(1, 10)]
    if any(v is not None for v in row_vals):
        print(f"Row {r:2d} | {str(row_vals[0]):<45} | {row_vals[1:]}")

wb.close()
