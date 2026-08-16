import openpyxl
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

excel_path = r"C:\Users\lap4all\Downloads\AOP_Hang_Nang_MAU_Vung_T7-T12_2026_final_v18.xlsx"

wb = openpyxl.load_workbook(excel_path, data_only=False)
sheet = wb['Kế hoạch T7']

print("=== IN CÔNG THỨC/GIÁ TRỊ CÁC Ô TÊN BƯU CỤC ===")
rows = [11, 21, 31, 41, 51, 61, 71, 81]
for r in rows:
    val = sheet.cell(row=r, column=1).value
    print(f"Cell A{r}: {val}")
wb.close()
