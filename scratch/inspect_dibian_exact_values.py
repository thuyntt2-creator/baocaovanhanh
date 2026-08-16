import openpyxl
import sys

sys.stdout.reconfigure(encoding='utf-8')

file_path = r"C:\Users\lap4all\Downloads\NTB_Input_Con_Thieu_Theo_Template_FLM_CRC.xlsx"
wb = openpyxl.load_workbook(file_path, data_only=True)
sheet = wb['Định biên & Sản lượng']

print("=== Exact Values in Định biên & Sản lượng rows 72-77 ===")
for r in range(72, 78):
    row_vals = [sheet.cell(r, c).value for c in range(1, 16)]
    print(f"Row {r}: {row_vals}")

wb.close()
