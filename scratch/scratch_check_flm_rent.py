import openpyxl
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

file_path = r"C:\Users\lap4all\Downloads\NTB_Input_FLM_CRC.xlsx"
wb = openpyxl.load_workbook(file_path, data_only=False)
sheet = wb['Chi phí FLM']

print("=== Rent Row in 'Chi phí FLM' sheet ===")
cols = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P']
# In Excel sheet Chi phí FLM, row 28 is likely the rent cost row.
# Let's print rows 25 to 35.
for r in range(25, 36):
    row_vals = []
    for c in cols:
        cell = sheet.cell(r, openpyxl.utils.column_index_from_string(c))
        val = cell.value
        if isinstance(val, str) and val.startswith('='):
            row_vals.append(f"F:{val}")
        else:
            row_vals.append(val)
    print(f"Row {r:2d} | {str(row_vals[0]):<35} | {row_vals[1:]}")

wb.close()
