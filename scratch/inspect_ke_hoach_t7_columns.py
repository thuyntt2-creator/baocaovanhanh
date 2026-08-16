import openpyxl
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

excel_path = r"C:\Users\lap4all\Downloads\AOP_Hang_Nang_MAU_Vung_T7-T12_2026_final_v18.xlsx"

wb = openpyxl.load_workbook(excel_path, data_only=True)
sheet = wb['Kế hoạch T7']

print("=== IN TẤT CẢ GIÁ TRỊ CỘT 1 TRONG KẾ HOẠCH T7 ===")
for r in range(1, 150):
    val = sheet.cell(row=r, column=1).value
    if val is not None:
        print(f"Row {r:03d}: {val}")
wb.close()
