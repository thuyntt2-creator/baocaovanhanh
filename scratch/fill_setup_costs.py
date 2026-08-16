import openpyxl
import win32com.client
import sys

sys.stdout.reconfigure(encoding='utf-8')

file_path = r"C:\Users\lap4all\Downloads\NTB_Input_Con_Thieu_Theo_Template_FLM_CRC.xlsx"

# ===== TÍNH TOÁN =====
# Mở mới per BCCK
mo_moi_per_bcck = 35_399_100 + 33_366_814 + 766_667 + 13_922_000  # = 83,454,581 đ
# Di dời per BCCK
di_doi_per_bcck = 7_043_685 + 20_215_215 + 700_000 + 13_922_000   # = 41,880,900 đ

# x4 cho 4 BCCK, chuyển sang triệu
mo_moi_total_trieu = mo_moi_per_bcck * 4 / 1_000_000   # = 333.818324 triệu
di_doi_total_trieu = di_doi_per_bcck * 4 / 1_000_000   # = 167.5236 triệu

print(f"D1. Setup mở mới (T7): {mo_moi_total_trieu:.6f} triệu")
print(f"D2. Di dời (T7): {di_doi_total_trieu:.6f} triệu")
print(f"Tháng T8-T12: 0.0 triệu")

# ===== ĐIỀN VÀO FILE =====
wb = openpyxl.load_workbook(file_path, data_only=False)
sheet = wb['NTB – Input']

cols_months = ['D', 'E', 'F', 'G', 'H', 'I']  # T7, T8, T9, T10, T11, T12

for idx, col in enumerate(cols_months):
    if idx == 0:  # T7 - tháng mở mới
        sheet[f"{col}38"] = mo_moi_total_trieu
        sheet[f"{col}39"] = di_doi_total_trieu
    else:  # T8-T12 - không phát sinh
        sheet[f"{col}38"] = 0.0
        sheet[f"{col}39"] = 0.0

wb.save(file_path)
wb.close()
print("\nĐã lưu file. Đang recalculate qua Excel COM...")

# Force recalculate
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
try:
    workbook = excel.Workbooks.Open(file_path)
    workbook.Save()
    workbook.Close()
    print("Recalculate và lưu hoàn tất!")
except Exception as e:
    print(f"Lỗi COM: {e}")
finally:
    excel.Quit()
