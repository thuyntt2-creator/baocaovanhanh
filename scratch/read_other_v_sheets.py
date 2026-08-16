import win32com.client
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
excel.AskToUpdateLinks = False

files_to_check = [
    r"C:\Users\lap4all\Downloads\AOP_Hang_Nang_MAU_Vung_T7-T12_2026_final_v16.xlsx",
    r"C:\Users\lap4all\Downloads\AOP_Hang_Nang_MAU_Vung_T7-T12_2026_final_v19.xlsx",
    r"C:\Users\lap4all\Downloads\AOP_Hang_Nang_MAU_Vung_T7-T12_2026_final_v20.xlsx",
]

for path in files_to_check:
    if not os.path.exists(path):
        print(f"File không tồn tại: {path}")
        continue
    try:
        wb = excel.Workbooks.Open(path, UpdateLinks=0, ReadOnly=True)
        print(f"\n=== FILE: {os.path.basename(path)} ===")
        sheet = wb.Sheets('Nguồn lực & chi phí')
        for r in range(1, 23):
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
        print(f"Lỗi đọc {os.path.basename(path)}: {e}")

excel.Quit()
