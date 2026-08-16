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
    
    # 1. Đọc sheet 0.3 Bưu cục Detail
    try:
        sheet_detail = wb.Sheets('0.3 Bưu cục Detail')
        print("=== SHEET: 0.3 Bưu Cục Detail (40 dòng đầu) ===")
        for r in range(1, 41):
            row_vals = [sheet_detail.Cells(r, c).Value for c in range(1, 15)]
            if any(row_vals):
                print(f"Row {r:02d}: {row_vals}")
    except Exception as e:
        print(f"Lỗi đọc 0.3 Bưu cục Detail: {e}")
        
    # 2. Đọc sheet Mặt bằng
    try:
        sheet_mb = wb.Sheets('Mặt bằng')
        print("\n=== SHEET: Mặt bằng (20 dòng đầu) ===")
        for r in range(1, 21):
            row_vals = [sheet_mb.Cells(r, c).Value for c in range(1, 15)]
            if any(row_vals):
                print(f"Row {r:02d}: {row_vals}")
    except Exception as e:
        print(f"Lỗi đọc Mặt bằng: {e}")
        
    wb.Close(False)
finally:
    excel.Quit()
