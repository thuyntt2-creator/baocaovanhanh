import openpyxl
import sys

sys.stdout.reconfigure(encoding='utf-8')

file_path = r"C:\Users\lap4all\Downloads\V3_AOP_Hang_NTB_T7-T12_2026 mới.xlsx"
wb = openpyxl.load_workbook(file_path, data_only=True)
sheet = wb['Forecast T7']

print("=== Forecast T7 Rows 8 to 15 ===")
for r in range(8, 16):
    row_vals = [sheet.cell(r, c).value for c in range(1, 15)]
    print(f"Row {r:2d}: {row_vals}")

wb.close()
