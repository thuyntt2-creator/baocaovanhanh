import openpyxl
import sys

sys.stdout.reconfigure(encoding='utf-8')

file_path = r"C:\Users\lap4all\Downloads\NTB_Input_Con_Thieu_Theo_Template_FLM_CRC.xlsx"
wb = openpyxl.load_workbook(file_path, data_only=False)
sheet = wb['NTB – Input']

print("Row 7 formula:", sheet['D7'].value)
print("Row 8 formula:", sheet['D8'].value)
print("Row 11 formula:", sheet['D11'].value)
wb.close()
