import openpyxl
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

excel_path = r"C:\Users\lap4all\Downloads\AOP_Hang_Nang_MAU_Vung_T7-T12_2026_final_v18.xlsx"

wb = openpyxl.load_workbook(excel_path, data_only=True)
sheet = wb['Tổng hợp xe']

print("=== IN DÒNG 5 ĐẾN 25 SHEET TỔNG HỢP XE ===")
for r in range(4, 25):
    row_vals = [sheet.cell(row=r+1, column=c+1).value for c in range(8)]
    row_str = " | ".join([str(v) if v is not None else "" for v in row_vals])
    print(f"Row {r+1:02d}: {row_str}")
wb.close()
