import win32com.client
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

excel_path = r"C:\Users\lap4all\Downloads\[NTB] So sánh topline H2 Mới - Cũ.xlsx"

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
        sheet = wb.Sheets(sname)
        print(f"\nSheet: {sheet.Name}")
        # In ra used range
        try:
            used_range = sheet.UsedRange
            vals = used_range.Value
            if vals is not None:
                # in 40 dòng đầu
                for r_idx, row in enumerate(vals[:40]):
                    if any(x is not None and x != "" for x in row):
                        formatted = []
                        for x in row[:12]:
                            if isinstance(x, (int, float)):
                                formatted.append(f"{x:,.2f}".rstrip('0').rstrip('.'))
                            elif x is None:
                                formatted.append('')
                            else:
                                formatted.append(str(x))
                        print(f"  Row {r_idx+1:02d}: {formatted}")
        except Exception as e:
            print(f"Lỗi đọc sheet {sname}: {e}")
            
    wb.Close(False)
except Exception as e:
    print(f"Lỗi: {e}")
finally:
    excel.Quit()
