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

targets = {
    'Chi phí xe T7': 715,
    'Chi phí NV giao T7': 690,
    'TỔNG CP T7': 1746,
    'Tổng nhân sự T7': 126,
    'Tổng NV giao T7': 76,
    'CP/đơn T7': 59401
}

try:
    wb = excel.Workbooks.Open(excel_path, UpdateLinks=0, ReadOnly=True)
    print(f"=== QUÉT TOÀN BỘ FILE: {os.path.basename(excel_path)} ===")
    for s_idx in range(1, wb.Sheets.Count + 1):
        sheet = wb.Sheets(s_idx)
        sname = sheet.Name
        try:
            used_range = sheet.UsedRange
            vals = used_range.Value
            if vals is None:
                continue
            
            # vals là tuple 2 chiều
            for r_idx, row in enumerate(vals):
                for c_idx, val in enumerate(row):
                    if val is not None:
                        for name, target in targets.items():
                            if (isinstance(val, (int, float)) and abs(val - target) < 1) or \
                               (isinstance(val, (int, float)) and abs(val/1e6 - target) < 1) or \
                               (isinstance(val, (int, float)) and abs(val/1e3 - target) < 1) or \
                               (isinstance(val, (int, float)) and abs(val - 59.401) < 0.1) or \
                               (isinstance(val, (int, float)) and abs(val - 59401) < 1):
                                print(f"[KHỚP V18] Sheet: {sname} | Ô {r_idx+1},{c_idx+1} ({sheet.Cells(r_idx+1, c_idx+1).Address}): Giá trị = {val} (Khớp với {name}: {target})")
        except Exception as e:
            print(f"Lỗi đọc sheet {sname}: {e}")
    wb.Close(False)
except Exception as e:
    print(f"Lỗi: {e}")
finally:
    excel.Quit()
