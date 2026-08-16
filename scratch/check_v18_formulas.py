import win32com.client
import sys
import os

sys.stdout.reconfigure(encoding='utf-8')

excel_path = r"C:\Users\lap4all\Downloads\AOP_Hang_Nang_MAU_Vung_T7-T12_2026_final_v18.xlsx"

if not os.path.exists(excel_path):
    print("File không tồn tại")
    sys.exit(1)

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
wb = excel.Workbooks.Open(excel_path)
sheet = wb.Sheets('Nguồn lực & chi phí')

print("=== IN DỮ LIỆU THỰC TẾ SHEET NGUỒN LỰC & CHI PHÍ ===")
# Đọc 120 dòng đầu, 8 cột đầu
for r in range(1, 105):
    row_vals = []
    row_formulas = []
    for c in range(1, 9):
        val = sheet.Cells(r, c).Value
        formula = sheet.Cells(r, c).Formula
        row_vals.append(val)
        row_formulas.append(formula)
        
    # Bỏ qua các dòng trống hoàn toàn
    if any(x is not None and x != "" for x in row_vals):
        # Format hiển thị giá trị
        formatted_vals = []
        for v in row_vals:
            if isinstance(v, float) or isinstance(v, int):
                formatted_vals.append(f"{v:,.1f}")
            else:
                formatted_vals.append(str(v))
        print(f"Row {r:03d} | Values  : {formatted_vals}")
        # print(f"Row {r:03d} | Formulas: {row_formulas}")

wb.Close(False)
excel.Quit()
