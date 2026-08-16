import openpyxl
import sys

sys.stdout.reconfigure(encoding='utf-8')

file_path = r"C:\Users\lap4all\Downloads\V3 AOP_NTB_T70-T12_2026.xlsx"
wb = openpyxl.load_workbook(file_path, data_only=True)
sheet = wb['Timeline tiếp nhận']

print("=== Sheet: Timeline tiếp nhận (first 20 rows) ===")
for r in range(1, 25):
    row_vals = [sheet.cell(r, c).value for c in range(1, 10)]
    if any(v is not None for v in row_vals):
        print(f"Row {r:2d}: {row_vals}")

wb.close()
