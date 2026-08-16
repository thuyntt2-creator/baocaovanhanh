import openpyxl
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

excel_path = r"C:\Users\lap4all\Downloads\AOP_Hang_NTB_T7-T12_2026_fixed_v12.xlsx"

if not os.path.exists(excel_path):
    print("File không tồn tại")
    sys.exit(1)

wb = openpyxl.load_workbook(excel_path, data_only=True)
sheet = wb['0.3 Bưu cục Detail']

print(f"=== ĐỌC SHEET 0.3 Bưu cục Detail trong {os.path.basename(excel_path)} ===")
for r in range(50, 100):
    row_vals = [sheet.cell(r, c).value for c in range(1, 15)]
    if any(x is not None for x in row_vals):
        print(f"Row {r:02d}: {[str(v) if v is not None else '' for v in row_vals]}")
