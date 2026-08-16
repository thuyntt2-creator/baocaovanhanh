import openpyxl
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

filepath = r"C:\Users\lap4all\Downloads\Telegram Desktop\7. NTB_2026.xlsx"
wb = openpyxl.load_workbook(filepath, data_only=False)
print("Sheets in 7. NTB_2026.xlsx:", wb.sheetnames)
wb.close()
