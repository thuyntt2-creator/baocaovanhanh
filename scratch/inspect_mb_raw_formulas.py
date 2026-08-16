import openpyxl
import sys

sys.stdout.reconfigure(encoding='utf-8')

file_path = r"C:\Users\lap4all\Downloads\V3 AOP_NTB_T70-T12_2026.xlsx"
wb = openpyxl.load_workbook(file_path, data_only=False)
sheet = wb['Mặt bằng']

print("=== Sheet: Mặt bằng (Raw Formulas & Values) ===")
for r in range(1, 20):
    row_vals = [sheet.cell(r, c).value for c in range(1, 15)]
    print(f"Row {r:2d}: {row_vals}")

wb.close()
