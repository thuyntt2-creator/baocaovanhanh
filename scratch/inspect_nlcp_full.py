import openpyxl
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

excel_path = r"C:\Users\lap4all\Downloads\AOP_Hang_Nang_MAU_Vung_T7-T12_2026_final_v18.xlsx"

wb = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)
sheet = wb['Nguồn lực & chi phí']

print("=== IN TOÀN BỘ SHEET NGUỒN LỰC & CHI PHÍ ===")
for r_idx in range(40):
    row_vals = []
    for c_idx in range(12):
        cell_val = sheet.cell(row=r_idx+1, column=c_idx+1).value
        row_vals.append(str(cell_val) if cell_val is not None else "")
    # Chỉ in nếu hàng không trống hoàn toàn
    if any(row_vals):
        print(f"Row {r_idx+1:02d}: {row_vals}")
wb.close()
