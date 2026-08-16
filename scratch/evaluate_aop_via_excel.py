import win32com.client
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

calculated_path = r"C:\Users\lap4all\Downloads\AOP_Hang_NTB_T7-T12_2026_calculated.xlsx"
abs_path = os.path.abspath(calculated_path)

if not os.path.exists(abs_path):
    print(f"File not found: {abs_path}")
    sys.exit(1)

print("Opening Excel...")
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
try:
    wb = excel.Workbooks.Open(abs_path)
    
    print("\n=== SHEET: Kênh & nhu cầu ===")
    sheet_k = wb.Sheets("Kênh & nhu cầu")
    for r in range(1, 11):
        row_vals = [sheet_k.Cells(r, c).Value for c in range(1, 8)]
        if any(v is not None for v in row_vals):
            print(f"Row {r:2d}: {row_vals}")
            
    print("\n=== SHEET: Nguồn lực & chi phí ===")
    sheet_n = wb.Sheets("Nguồn lực & chi phí")
    for r in range(1, 19):
        row_vals = [sheet_n.Cells(r, c).Value for c in range(1, 8)]
        if any(v is not None for v in row_vals):
            print(f"Row {r:2d}: {row_vals}")

    print("\n=== SHEET: Mặt bằng ===")
    sheet_m = wb.Sheets("Mặt bằng")
    for r in range(1, 20):
        row_vals = [sheet_m.Cells(r, c).Value for c in range(1, 14)]
        if any(v is not None for v in row_vals):
            print(f"Row {r:2d}: {row_vals}")

    wb.Close(SaveChanges=False)
except Exception as e:
    print(f"Error: {e}")
finally:
    excel.Quit()

print("\nExcel closed.")

