import openpyxl
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

excel_path = r"C:\Users\lap4all\Downloads\AOP_Hang_Nang_MAU_Vung_T7-T12_2026_final_v18.xlsx"

wb = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)
sheet = wb['Nguồn lực & chi phí']

print("=== ĐỌC DÒNG 23-35 SHEET NGUỒN LỰC & CHI PHÍ ===")
for r_idx in range(22, 35):
    row_vals = [cell.value for cell in sheet[r_idx+1]]
    print(f"Row {r_idx+1:02d}: {row_vals}")
wb.close()
