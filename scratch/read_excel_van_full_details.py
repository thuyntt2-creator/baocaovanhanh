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
    
    # 1. Đọc sheet Nguồn lực & chi phí từ hàng 20 đến 100
    try:
        sheet_nl = wb.Sheets('Nguồn lực & chi phí')
        print("=== SHEET: NGUỒN LỰC & CHI PHÍ (Hàng 20-100) ===")
        for r in range(20, 101):
            row_vals = [sheet_nl.Cells(r, c).Value for c in range(1, 10)]
            if any(row_vals):
                print(f"Row {r:03d}: {row_vals}")
    except Exception as e:
        print(f"Lỗi đọc sheet: {e}")
        
    wb.Close(False)
finally:
    excel.Quit()
