import openpyxl
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

filepath = r"C:\Users\lap4all\Downloads\NTB_Input_Da_Dien_FLM_CRC_2.xlsx"
wb = openpyxl.load_workbook(filepath, data_only=True)
sheet = wb['NTB – Input']

cols = ['D', 'E', 'F', 'G', 'H', 'I']
print("=== Row 25 (Số xe cần có) in NTB_Input_Da_Dien_FLM_CRC_2.xlsx ===")
print([sheet[f"{c}25"].value for c in cols])
wb.close()
