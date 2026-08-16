import openpyxl
import sys

sys.stdout.reconfigure(encoding='utf-8')

file_path = r"C:\Users\lap4all\Downloads\NTB_Input_Con_Thieu_Theo_Template_FLM_CRC.xlsx"
wb = openpyxl.load_workbook(file_path, data_only=False)
sheet = wb['Chi phí FLM']

print("=== Col headers in 'Chi phí FLM' ===")
for r in [4, 5, 6]:
    row_vals = [sheet.cell(r, c).value for c in range(1, 18)]
    print(f"Row {r:2d}: {row_vals}")
wb.close()
