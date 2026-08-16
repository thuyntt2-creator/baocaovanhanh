import win32com.client
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

downloads_dir = r"C:\Users\lap4all\Downloads"

excel_files = []
for f in os.listdir(downloads_dir):
    if f.endswith('.xlsx') and not f.startswith('~$'):
        excel_files.append(os.path.join(downloads_dir, f))

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
excel.AskToUpdateLinks = False

targets = {
    'Tổng CP T9': 2408.6,
    'Tổng CP T10': 2522.1,
    'Tổng CP T11': 2803.9,
    'Tổng CP T12': 2914.5
}

try:
    for path in excel_files:
        try:
            wb = excel.Workbooks.Open(path, UpdateLinks=0, ReadOnly=True)
            for s_idx in range(1, wb.Sheets.Count + 1):
                sheet = wb.Sheets(s_idx)
                sname = sheet.Name
                try:
                    used_range = sheet.UsedRange
                    vals = used_range.Value
                    if vals is None:
                        continue
                    
                    # Quét mảng 2 chiều
                    for r_idx, row in enumerate(vals):
                        for c_idx, val in enumerate(row):
                            if val is not None and isinstance(val, (int, float)):
                                # Kiểm tra xem có khớp với target nào không (chia cho 1e6 hoặc trực tiếp)
                                for name, target in targets.items():
                                    if abs(val - target) < 0.1 or abs(val/1e6 - target) < 0.1:
                                        print(f"[KHỚP] File: {os.path.basename(path)} | Sheet: {sname} | Ô {r_idx+1},{c_idx+1} ({sheet.Cells(r_idx+1, c_idx+1).Address}): Giá trị = {val} ({name})")
                except Exception as e:
                    pass
            wb.Close(False)
        except Exception as e:
            print(f"Lỗi mở {os.path.basename(path)}: {e}")
finally:
    excel.Quit()
