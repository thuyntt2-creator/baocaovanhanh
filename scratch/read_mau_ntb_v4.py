import win32com.client
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
excel.AskToUpdateLinks = False

files_to_check = [
    r"C:\Users\lap4all\Downloads\AOP_MAU_NTB_T7-T12_2026_v4.xlsx",
    r"C:\Users\lap4all\Downloads\AOP_MAU_NTB_Cach3_TBDeu (1).xlsx",
]

for path in files_to_check:
    if not os.path.exists(path):
        print(f"File không tồn tại: {path}")
        continue
    try:
        wb = excel.Workbooks.Open(path, UpdateLinks=0, ReadOnly=True)
        print(f"\n=== FILE: {os.path.basename(path)} ===")
        # In ra tất cả sheet để xem
        sheet_names = [wb.Sheets(i).Name for i in range(1, wb.Sheets.Count + 1)]
        print("Sheets:", sheet_names)
        
        # Tìm sheet Nguồn lực & chi phí hoặc sheet tương đương
        nlcp_sheet = None
        for sname in sheet_names:
            if 'nguồn lực' in sname.lower() or 'nlcp' in sname.lower():
                nlcp_sheet = wb.Sheets(sname)
                break
                
        if nlcp_sheet:
            print(f"Đọc sheet: {nlcp_sheet.Name}")
            for r in range(1, 30):
                row_vals = []
                for c in range(1, 9):
                    row_vals.append(nlcp_sheet.Cells(r, c).Value)
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
