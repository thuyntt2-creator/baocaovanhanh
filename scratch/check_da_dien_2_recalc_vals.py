import openpyxl
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

filepath = r"C:\Users\lap4all\Downloads\NTB_Input_Da_Dien_FLM_CRC_2.xlsx"
wb = openpyxl.load_workbook(filepath, data_only=True)
sheet = wb['Chi phí FLM']

print("=== NTB_Input_Da_Dien_FLM_CRC_2.xlsx - Chi phí FLM J36, J37, J38 ===")
print("J36 (Total):", sheet['J36'].value)
print("J37 (Giao):", sheet['J37'].value)
print("J38 (Lấy):", sheet['J38'].value)

sheet_input = wb['NTB – Input']
print("\nNTB – Input Row 41 col D formula:", wb['NTB – Input']['D41'].value)
print("NTB – Input Row 41 col D value:", sheet_input['D41'].value)
wb.close()
