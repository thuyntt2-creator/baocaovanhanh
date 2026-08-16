import win32com.client
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

excel_path = r"C:\Users\lap4all\Downloads\AOP_Hang_Nang_MAU_Vung_T7-T12_2026_final_v18.xlsx"

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
excel.AskToUpdateLinks = False

try:
    wb = excel.Workbooks.Open(excel_path, UpdateLinks=0, ReadOnly=True)
    sheet = wb.Sheets('Kế hoạch T7')
    print("=== TÌM MAX XE XẾP THỰC TẾ TỪNG BCCK TRONG T7 ===")
    
    # Các dòng xe xếp thực tế của 4 bưu cục:
    # Nha Trang: dòng 18
    # Đơn Dương: dòng 28
    # Di Linh: dòng 38
    # Đức Linh: dòng 48
    # Tổng: dòng 94
    
    hubs = {
        'Nha Trang (Row 18)': 18,
        'Đơn Dương (Row 28)': 28,
        'Di Linh (Row 38)': 38,
        'Đức Linh (Row 48)': 48,
        'TỔNG CẢ VÙNG (Row 94)': 94
    }
    
    for name, row in hubs.items():
        vals = []
        # Cột B đến AF (cột 2 đến 32)
        for c in range(2, 33):
            val = sheet.Cells(row, c).Value
            if val is not None:
                vals.append(val)
        if vals:
            print(f"{name}: Average = {sum(vals)/len(vals):.2f} | Max = {max(vals)} | Min = {min(vals)}")
            
    wb.Close(False)
except Exception as e:
    print(f"Lỗi: {e}")
finally:
    excel.Quit()
