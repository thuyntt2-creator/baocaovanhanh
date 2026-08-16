import openpyxl
import sys

sys.stdout.reconfigure(encoding='utf-8')

file_path = r"C:\Users\lap4all\Downloads\NTB_Input_Con_Thieu_Theo_Template_FLM_CRC.xlsx"
wb = openpyxl.load_workbook(file_path, data_only=True)
sheet = wb['Định biên & Sản lượng']

print("=== Sheet: Định biên & Sản lượng (Rows 50 to 90) ===")
for r in range(50, 95):
    row_vals = [sheet.cell(r, c).value for c in range(1, 16)]
    if any(v is not None for v in row_vals):
        print(f"Row {r:2d}: {row_vals}")

wb.close()
