import openpyxl
import sys

sys.stdout.reconfigure(encoding='utf-8')

file_path = r"C:\Users\lap4all\Downloads\NTB_Input_Con_Thieu_Theo_Template_FLM_CRC.xlsx"
wb = openpyxl.load_workbook(file_path, data_only=True)
sheet = wb['Định biên & Sản lượng']

print("=== All rows in Định biên & Sản lượng ===")
for r in range(1, 95):
    row_vals = [sheet.cell(r, c).value for c in range(1, 16)]
    # Check if there are any non-empty values
    if any(v is not None for v in row_vals):
        # Format values to avoid showing long lists of Nones
        non_empty = [f"Col{c}:{v}" for c, v in enumerate(row_vals, 1) if v is not None]
        print(f"Row {r:2d}: {non_empty}")

wb.close()
