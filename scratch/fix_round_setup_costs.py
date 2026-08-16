import openpyxl
import win32com.client
import sys

sys.stdout.reconfigure(encoding='utf-8')

file_path = r"C:\Users\lap4all\Downloads\NTB_Input_Con_Thieu_Theo_Template_FLM_CRC.xlsx"

# Làm tròn về 2 chữ số thập phân
mo_moi_rounded = round(333.818324, 2)   # = 333.82
di_doi_rounded  = round(83.7618, 2)     # = 83.76

print(f"D1. Setup mở mới → {mo_moi_rounded} triệu")
print(f"D2. Di dời       → {di_doi_rounded} triệu")

wb = openpyxl.load_workbook(file_path, data_only=False)
sheet = wb['NTB – Input']

cols = ['D', 'E', 'F', 'G', 'H', 'I']
for col in cols:
    sheet[f"{col}38"] = mo_moi_rounded
    sheet[f"{col}39"] = di_doi_rounded

wb.save(file_path)
wb.close()
print("Đã lưu. Đang recalculate qua Excel COM...")

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
try:
    workbook = excel.Workbooks.Open(file_path)
    workbook.Save()
    workbook.Close()
    print("Hoàn tất!")
except Exception as e:
    print(f"Lỗi COM: {e}")
finally:
    excel.Quit()

# Xác nhận lại
wb2 = openpyxl.load_workbook(file_path, data_only=True)
s = wb2['NTB – Input']
print(f"\nRow 38: {[s[f'{c}38'].value for c in cols]}")
print(f"Row 39: {[s[f'{c}39'].value for c in cols]}")
wb2.close()
