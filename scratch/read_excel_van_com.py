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
    
    print("=== TÊN CÁC SHEET TRONG V2_AOP_van_fixed_final.xlsx ===")
    for s_idx in range(1, wb.Sheets.Count + 1):
        print(f"- {wb.Sheets(s_idx).Name}")
        
    # Thử đọc sheet Nguồn lực & chi phí
    try:
        sheet_nl = wb.Sheets('Nguồn lực & chi phí')
        print("\n=== SHEET: NGUỒN LỰC & CHI PHÍ ===")
        for r in range(1, 26):
            row_vals = [sheet_nl.Cells(r, c).Value for c in range(1, 8)]
            print(f"Row {r:02d}: {row_vals}")
    except Exception as e:
        print(f"Không thể đọc sheet Nguồn lực & chi phí: {e}")
        
    # Thử đọc sheet Tổng hợp xe
    try:
        sheet_th = wb.Sheets('Tổng hợp xe')
        print("\n=== SHEET: TỔNG HỢP XE ===")
        for r in range(1, 15):
            row_vals = [sheet_th.Cells(r, c).Value for c in range(1, 8)]
            print(f"Row {r:02d}: {row_vals}")
    except Exception as e:
        print(f"Không thể đọc sheet Tổng hợp xe: {e}")
        
    wb.Close(False)
finally:
    excel.Quit()
