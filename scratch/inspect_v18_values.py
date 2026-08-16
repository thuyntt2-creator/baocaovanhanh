import openpyxl
import sys

sys.stdout.reconfigure(encoding='utf-8')

path = r"C:\Users\lap4all\Downloads\AOP_Hang_Nang_MAU_Vung_T7-T12_2026_final_v18.xlsx"
wb = openpyxl.load_workbook(path, data_only=True)
sheet = wb['Nguồn lực & chi phí']

print(f"=== Sheet: {sheet.title} trong file {path} ===")
for r in range(1, 23):
    row_vals = [sheet.cell(r, c).value for c in range(1, 9)]
    print(f"Row {r}: {row_vals}")
