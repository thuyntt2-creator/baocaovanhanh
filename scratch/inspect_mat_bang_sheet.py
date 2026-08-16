import openpyxl
import sys

sys.stdout.reconfigure(encoding='utf-8')
file_path = r"C:\Users\lap4all\Downloads\AOP_Hang_Nang_MAU_Vung_T7-T12_2026_final_v18.xlsx"
wb = openpyxl.load_workbook(file_path, read_only=True)
print("Sheet names in workbook:")
for idx, name in enumerate(wb.sheetnames):
    print(f"  Index {idx}: {repr(name)}")
