import win32com.client
import os
import sys
import glob

sys.stdout.reconfigure(encoding='utf-8')

downloads_dir = r"C:\Users\lap4all\Downloads"
all_files = glob.glob(os.path.join(downloads_dir, "*.xlsx"))

# Lọc các file có tên chứa AOP hoặc NTB và dung lượng < 10MB
excel_files = []
for p in all_files:
    fname = os.path.basename(p)
    if fname.startswith("~$"):
        continue
    size_mb = os.path.getsize(p) / (1024 * 1024)
    name_lower = fname.lower()
    if ("aop" in name_lower or "ntb" in name_lower) and size_mb < 10:
        excel_files.append(p)

print(f"=== ĐANG DÙNG COM QUÉT {len(excel_files)} FILE EXCEL IN DOWNLOADS ===")

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False

targets = {
    'Chi phí xe T7': 715,
    'Chi phí NV giao T7': 690,
    'TỔNG CP T7': 1746,
    'Tổng nhân sự T7': 126,
    'Tổng NV giao T7': 76,
    'CP/đơn T7': 59401
}

for path in excel_files:
    fname = os.path.basename(path)
    try:
        wb = excel.Workbooks.Open(path)
    except Exception as e:
        # print(f"Không thể mở {fname}: {e}")
        continue
        
    for s_idx in range(1, wb.Sheets.Count + 1):
        sheet = wb.Sheets(s_idx)
        sname = sheet.Name
        # Quét vùng A1:Z100
        for r in range(1, 101):
            for c in range(1, 27):
                val = sheet.Cells(r, c).Value
                if val is not None:
                    try:
                        # Kiểm tra khớp target
                        for name, target in targets.items():
                            if (isinstance(val, (int, float)) and abs(val - target) < 1) or \
                               (isinstance(val, (int, float)) and abs(val/1e6 - target) < 1) or \
                               (isinstance(val, (int, float)) and abs(val/1e3 - target) < 1) or \
                               (isinstance(val, (int, float)) and abs(val - 59.401) < 0.1) or \
                               (isinstance(val, (int, float)) and abs(val - 59401) < 1):
                                print(f"[KHỚP COM] File: {fname} | Sheet: {sname} | Ô {r},{c} ({sheet.Cells(r, c).Address}): Giá trị = {val} (Khớp với {name}: {target})")
                    except:
                        pass
    wb.Close(False)

excel.Quit()
print("=== QUÉT HOÀN TẤT ===")
