import openpyxl
import sys

sys.stdout.reconfigure(encoding='utf-8')

file_path = r"C:\Users\lap4all\Downloads\NTB_Input_Con_Thieu_Theo_Template_FLM_CRC.xlsx"
wb = openpyxl.load_workbook(file_path, data_only=False)
sheet = wb['Định biên & Sản lượng']

print("=== Định biên & Sản lượng row 17 formulas ===")
cols = ['I', 'J', 'K', 'L', 'M', 'N']
for c in cols:
    print(f"  {c}17: {sheet[f'{c}17'].value}")
wb.close()
