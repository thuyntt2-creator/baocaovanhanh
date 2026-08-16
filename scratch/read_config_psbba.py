import win32com.client
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

excel_path = r"C:\Users\lap4all\Downloads\config_psbba_NTB.xlsx"

if not os.path.exists(excel_path):
    print("File không tồn tại")
    sys.exit(1)

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
excel.AskToUpdateLinks = False

try:
    wb = excel.Workbooks.Open(excel_path, UpdateLinks=0, ReadOnly=True)
    sheet_names = [wb.Sheets(i).Name for i in range(1, wb.Sheets.Count + 1)]
    print(f"=== FILE: {os.path.basename(excel_path)} (Sheets: {sheet_names}) ===")
    
    for sname in sheet_names:
        if 'topline' in sname.lower() or 'nguồn lực' in sname.lower():
            sheet = wb.Sheets(sname)
            print(f"\nSheet: {sheet.Name}")
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
                    print(f"  Row {r:02d}: {formatted}")
    wb.Close(False)
except Exception as e:
    print(f"Lỗi: {e}")
finally:
    excel.Quit()
