import openpyxl
import win32com.client
import sys

sys.stdout.reconfigure(encoding='utf-8')

file_path = r"C:\Users\lap4all\Downloads\NTB_Input_Con_Thieu_Theo_Template_FLM_CRC.xlsx"

wb = openpyxl.load_workbook(file_path, data_only=False)
sheet = wb['NTB – Input']

cols = ['D', 'E', 'F', 'G', 'H', 'I']

print("Cập nhật Row 43 (TỔNG CHI PHÍ FLM) = Row41 + Row42...")
for col in cols:
    sheet[f"{col}43"] = f"={col}41+{col}42"

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
    print(f"Lỗi: {e}")
finally:
    excel.Quit()

# Verify
wb2 = openpyxl.load_workbook(file_path, data_only=True)
s = wb2['NTB – Input']
months = ['T7','T8','T9','T10','T11','T12']
print("\n=== Kết quả sau khi sửa (triệu đ) ===")
for r, lbl in [(41,"TỔNG CHI PHÍ GIAO"), (42,"TỔNG CHI PHÍ NHẬN"), (43,"TỔNG CHI PHÍ FLM"), (44,"Chi phí / đơn"), (45,"Tổng cost / kg")]:
    vals = [s[f"{c}{r}"].value for c in cols]
    if vals[0] is not None:
        formatted = [f"{v:,.0f}" if v and abs(v) > 100 else f"{v:,.2f}" if v else "–" for v in vals]
    else:
        formatted = ["–"] * 6
    print(f"  Row {r:2d} ({lbl}): {formatted}")
wb2.close()
