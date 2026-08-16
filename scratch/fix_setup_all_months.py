import openpyxl
import win32com.client
import sys

sys.stdout.reconfigure(encoding='utf-8')

file_path = r"C:\Users\lap4all\Downloads\NTB_Input_Con_Thieu_Theo_Template_FLM_CRC.xlsx"

mo_moi_per_bcck = 35_399_100 + 33_366_814 + 766_667 + 13_922_000   # = 83,454,581 đ
di_doi_per_bcck = 7_043_685 + 20_215_215 + 700_000 + 13_922_000    # = 41,880,900 đ

mo_moi_trieu = mo_moi_per_bcck * 4 / 1_000_000   # 4 BC mở mới -> 333.818 triệu/tháng
di_doi_trieu = di_doi_per_bcck * 2 / 1_000_000   # 2 BC di dời -> 83.762 triệu/tháng

print(f"D1. Setup mở mới (4 BC): {mo_moi_trieu:.4f} triệu/tháng")
print(f"D2. Di dời (2 BC): {di_doi_trieu:.4f} triệu/tháng")

wb = openpyxl.load_workbook(file_path, data_only=False)
sheet = wb['NTB – Input']

cols = ['D', 'E', 'F', 'G', 'H', 'I']  # T7 → T12

for col in cols:
    sheet[f"{col}38"] = mo_moi_trieu
    sheet[f"{col}39"] = di_doi_trieu

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
r38 = [s[f"{c}38"].value for c in cols]
r39 = [s[f"{c}39"].value for c in cols]
print(f"\nDòng 38 (Setup): {r38}")
print(f"Dòng 39 (Di dời): {r39}")
wb2.close()
