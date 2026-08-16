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
wb = excel.Workbooks.Open(excel_path)

targets = {
    'Chi phí xe T7': 715,
    'Chi phí NV giao T7': 690,
    'TỔNG CP T7': 1746,
    'Tổng nhân sự T7': 126,
    'Tổng NV giao T7': 76,
}

print(f"=== Tìm kiếm bằng COM trong file: {os.path.basename(excel_path)} ===")

for s_idx in range(1, wb.Sheets.Count + 1):
    sheet = wb.Sheets(s_idx)
    sname = sheet.Name
    # Đọc vùng dữ liệu A1:Z100
    for r in range(1, 101):
        for c in range(1, 27):
            val = sheet.Cells(r, c).Value
            if val is not None:
                try:
                    # Kiểm tra xem có khớp giá trị mục tiêu không
                    for name, target in targets.items():
                        if (isinstance(val, (int, float)) and abs(val - target) < 1) or \
                           (isinstance(val, (int, float)) and abs(val/1e6 - target) < 1) or \
                           (isinstance(val, (int, float)) and abs(val/1e3 - target) < 1):
                            print(f"[{sname}] Ô {r},{c} ({sheet.Cells(r, c).Address}): Giá trị = {val} (Khớp với {name}: {target})")
                except:
                    pass

wb.Close(False)
excel.Quit()
