import win32com.client
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

excel_path = r"C:\Users\lap4all\Downloads\AOP_Hang_Nang_MAU_Vung_T7-T12_2026_final_v18.xlsx"

if not os.path.exists(excel_path):
    print("File không tồn tại")
    sys.exit(1)

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
excel.AskToUpdateLinks = False

try:
    wb = excel.Workbooks.Open(excel_path, UpdateLinks=0, ReadOnly=True)
    
    # 1. Đọc sheet Tổng hợp xe
    sheet_th = wb.Sheets('Tổng hợp xe')
    print("=== SHEET: TỔNG HỢP XE ===")
    for r in range(1, 15):
        row_vals = [sheet_th.Cells(r, c).Value for c in range(1, 8)]
        print(f"Row {r:02d}: {row_vals}")
        
    # 2. Đọc sheet Nguồn lực & chi phí
    sheet_nl = wb.Sheets('Nguồn lực & chi phí')
    print("\n=== SHEET: NGUỒN LỰC & CHI PHÍ (25 dòng đầu) ===")
    for r in range(1, 26):
        row_vals = [sheet_nl.Cells(r, c).Value for c in range(1, 8)]
        print(f"Row {r:02d}: {row_vals}")
        
    wb.Close(False)
finally:
    excel.Quit()
