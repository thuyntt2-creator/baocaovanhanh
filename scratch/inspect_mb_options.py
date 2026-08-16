import openpyxl
import sys

sys.stdout.reconfigure(encoding='utf-8')

file_path = r"C:\Users\lap4all\Downloads\V3 AOP_NTB_T70-T12_2026.xlsx"
wb = openpyxl.load_workbook(file_path, data_only=True)
sheet = wb['Mặt bằng']

print("=== Column H (Tình trạng) values ===")
for r in range(3, 10):
    print(f"Row {r}: {sheet.cell(r, 8).value}")

wb.close()
