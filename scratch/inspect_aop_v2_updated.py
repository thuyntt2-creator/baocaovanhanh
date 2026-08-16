import openpyxl
import sys
import os

sys.stdout.reconfigure(encoding='utf-8')

path = r"C:\Users\lap4all\Downloads\AOP_V2_updated_2.xlsx"

if not os.path.exists(path):
    print("File không tồn tại")
    sys.exit(1)

wb = openpyxl.load_workbook(path, data_only=True)
sheet = wb['Nguồn lực & chi phí']

print(f"=== Sheet: {sheet.title} trong file {path} ===")
for r in range(1, 40):
    row_vals = [sheet.cell(r, c).value for c in range(1, 9)]
    if any(x is not None for x in row_vals):
        # format float numbers
        formatted = []
        for x in row_vals:
            if isinstance(x, (int, float)):
                formatted.append(f"{x:,.2f}".rstrip('0').rstrip('.'))
            else:
                formatted.append(str(x))
        print(f"Row {r:02d}: {formatted}")
