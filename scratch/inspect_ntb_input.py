import openpyxl
import sys

sys.stdout.reconfigure(encoding='utf-8')

file_path = r"C:\Users\lap4all\Downloads\NTB_Input_Con_Thieu_Theo_Template_FLM_CRC.xlsx"
wb = openpyxl.load_workbook(file_path, data_only=False)
sheet = wb['NTB – Input']

print("=== Row details for 'NTB – Input' ===")
for r in range(1, 48):
    row_vals = []
    for c in range(1, 11):
        val = sheet.cell(r, c).value
        # If it's a merged cell and we load without data_only, it might be None or a value.
        # Let's print the actual value or formula.
        row_vals.append(val)
    print(f"Row {r:2d}: {row_vals}")

wb.close()
