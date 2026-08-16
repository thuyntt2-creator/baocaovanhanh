import win32com.client
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

excel_path = r"C:\Users\lap4all\Downloads\V2_AOP_van_fixed_final.xlsx"

if not os.path.exists(excel_path):
    print("File không tồn tại")
    sys.exit(1)

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
excel.AskToUpdateLinks = False

try:
    wb = excel.Workbooks.Open(excel_path, UpdateLinks=0, ReadOnly=True)
    sheet = wb.Sheets('0.3 Bưu cục Detail')
    
    print("=== TÌM KIẾM CHI TIẾT CÁC BƯU CỤC TRONG SHEET 0.3 BƯU CỤC DETAIL ===")
    
    # In ra tiêu đề (dòng 1, 2)
    print(f"Row 01: {[sheet.Cells(1, c).Value for c in range(1, 15)]}")
    print(f"Row 02: {[sheet.Cells(2, c).Value for c in range(1, 15)]}")
    
    keywords = ["nha trang", "di linh", "đơn dương", "đức linh"]
    for r in range(3, 100):
        row_vals = [sheet.Cells(r, c).Value for c in range(1, 15)]
        row_str = " | ".join([str(v) for v in row_vals if v is not None])
        if any(kw in row_str.lower() for kw in keywords):
            print(f"Row {r:03d}: {row_vals}")
            
    wb.Close(False)
finally:
    excel.Quit()
