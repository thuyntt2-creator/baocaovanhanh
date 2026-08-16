import openpyxl
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

file_path = r"C:\Users\lap4all\Downloads\NTB_Input_FLM_CRC.xlsx"
wb = openpyxl.load_workbook(file_path, data_only=False)
sheet = wb['NTB – Input']

print("=== Formulas for Rows 41 and 42 in 'NTB – Input' ===")
cols = ['D', 'E', 'F', 'G', 'H', 'I']
for col in cols:
    print(f"Cell {col}41: {sheet[f'{col}41'].value}")
    print(f"Cell {col}42: {sheet[f'{col}42'].value}")

wb.close()
