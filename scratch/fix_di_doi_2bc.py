import openpyxl
import win32com.client
import sys

sys.stdout.reconfigure(encoding='utf-8')

file_path = r"C:\Users\lap4all\Downloads\NTB_Input_Con_Thieu_Theo_Template_FLM_CRC.xlsx"

di_doi_per_bcck = 7_043_685 + 20_215_215 + 700_000 + 13_922_000   # = 41,880,900 đ
di_doi_2bc_trieu = di_doi_per_bcck * 2 / 1_000_000   # = 83.7618 triệu

print(f"D2. Di dời (2 BC): {di_doi_2bc_trieu:.4f} triệu")

wb = openpyxl.load_workbook(file_path, data_only=False)
sheet = wb['NTB – Input']

# Chỉ cập nhật T7 (col D) dòng 39 - Di dời
sheet["D39"] = di_doi_2bc_trieu
# T8-T12 giữ nguyên 0

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

# Xác nhận lại giá trị
wb2 = openpyxl.load_workbook(file_path, data_only=True)
s = wb2['NTB – Input']
print(f"\nKiểm tra dòng 38 (Setup): T7={s['D38'].value:.4f} triệu")
print(f"Kiểm tra dòng 39 (Di dời): T7={s['D39'].value:.4f} triệu")
wb2.close()
