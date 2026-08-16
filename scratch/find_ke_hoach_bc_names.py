import openpyxl
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

excel_path = r"C:\Users\lap4all\Downloads\AOP_Hang_Nang_MAU_Vung_T7-T12_2026_final_v18.xlsx"

wb = openpyxl.load_workbook(excel_path, data_only=True)
sheet = wb['Kế hoạch T7']

print("=== IN CÁC DÒNG TIÊU ĐỀ BLOCK BƯU CỤC IN SHEET KẾ HOẠCH T7 ===")
for r in [11, 21, 31, 41, 51, 61, 71, 81]:
    vals = [sheet.cell(row=r, column=c).value for c in range(1, 15)]
    print(f"Row {r:03d}: {vals}")
wb.close()
