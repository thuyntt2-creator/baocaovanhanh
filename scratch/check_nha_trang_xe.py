import win32com.client
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

excel_path = r"C:\Users\lap4all\Downloads\AOP_Hang_Nang_MAU_Vung_T7-T12_2026_final_v18.xlsx"

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
excel.AskToUpdateLinks = False

months = ['T7', 'T8', 'T9', 'T10', 'T11', 'T12']

try:
    wb = excel.Workbooks.Open(excel_path, UpdateLinks=0, ReadOnly=True)
    print("=== TRUNG BÌNH XE NHA TRANG (Dòng 18) ===")
    for m in months:
        sheet = wb.Sheets(f'Kế hoạch {m}')
        end_col = 32 if m in ['T7', 'T8', 'T10', 'T12'] else 31
        vals = []
        for c in range(2, end_col + 1):
            val = sheet.Cells(18, c).Value
            if val is not None:
                vals.append(val)
        if vals:
            print(f"Tháng {m}: Average = {sum(vals)/len(vals):.2f} | Max = {max(vals)}")
            
    wb.Close(False)
except Exception as e:
    print(f"Lỗi: {e}")
finally:
    excel.Quit()
