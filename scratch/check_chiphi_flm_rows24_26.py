import openpyxl
import sys

sys.stdout.reconfigure(encoding='utf-8')

file_path = r"C:\Users\lap4all\Downloads\NTB_Input_Da_Dien_FLM_CRC_2.xlsx"
wb = openpyxl.load_workbook(file_path, data_only=False)
sheet = wb['Chi phí FLM']

print("=== Chi phí FLM rows 24-26 formulas ===")
for r in [24, 25, 26]:
    print(f"Row {r:2d} | D: {sheet[f'D{r}'].value}")
wb.close()
