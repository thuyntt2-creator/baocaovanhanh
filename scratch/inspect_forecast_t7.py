import openpyxl
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

excel_path = r"C:\Users\lap4all\Downloads\AOP_Hang_Nang_MAU_Vung_T7-T12_2026_final_v18.xlsx"

wb = openpyxl.load_workbook(excel_path, data_only=True)
sheet = wb['Forecast T7']

print("=== IN GIÁ TRỊ TÊN BƯU CỤC TỪ SHEET FORECAST T7 ===")
for r in range(4, 12):
    val = sheet.cell(row=r, column=1).value
    print(f"Row {r} (Col A): {val}")
wb.close()
