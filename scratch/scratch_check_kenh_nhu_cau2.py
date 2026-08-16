import openpyxl
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

file_path = r"C:\Users\lap4all\Downloads\AOP_Hang_NTB_T7-T12_2026_calculated.xlsx"
wb = openpyxl.load_workbook(file_path, data_only=True)
sheet = wb['Kênh & nhu cầu']

print("=== Values in 'Kênh & nhu cầu' (Rows 1-7) ===")
for r in range(1, 8):
    row_vals = [sheet.cell(r, c).value for c in range(1, 10)]
    if any(v is not None for v in row_vals):
        print(f"Row {r:2d}: {row_vals}")

wb.close()
