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
    sheet = wb.Sheets('1. Thông số')
    print("=== ĐỌC SHEET 1. Thông số ===")
    for r in range(1, 40):
        row_vals = []
        for c in range(1, 9):
            row_vals.append(sheet.Cells(r, c).Value)
        if any(x is not None for x in row_vals):
            formatted = []
            for x in row_vals:
                if isinstance(x, (int, float)):
                    formatted.append(f"{x:,.2f}".rstrip('0').rstrip('.'))
                else:
                    formatted.append(str(x))
            print(f"Row {r:02d}: {formatted}")
    wb.Close(False)
except Exception as e:
    print(f"Lỗi: {e}")
finally:
    excel.Quit()
