import openpyxl
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

file_path = r"C:\Users\lap4all\Downloads\AOP_Hang_NTB_T7-T12_2026_calculated.xlsx"
wb = openpyxl.load_workbook(file_path, data_only=False)
sheet = wb['Nguồn lực & chi phí']

print("=== Formulas for Row 6 in 'Nguồn lực & chi phí' ===")
cols = ['B', 'C', 'D', 'E', 'F', 'G']
for col in cols:
    cell = sheet[f"{col}6"]
    print(f"Cell {col}6: {cell.value}")

wb.close()
