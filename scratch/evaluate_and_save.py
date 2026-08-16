import win32com.client
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

calculated_path = r"C:\Users\lap4all\Downloads\AOP_Hang_NTB_T7-T12_2026_calculated.xlsx"
abs_path = os.path.abspath(calculated_path)

if not os.path.exists(abs_path):
    print(f"❌ File not found: {abs_path}")
    sys.exit(1)

print("🚀 Opening Excel Application to calculate formulas...")
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
try:
    wb = excel.Workbooks.Open(abs_path)
    # Force full calculation
    excel.CalculateFull()
    print("💾 Saving calculated workbook...")
    wb.Save()
    wb.Close(SaveChanges=True)
    print("✅ Calculations evaluated and saved successfully!")
except Exception as e:
    print(f"❌ Error during Excel calculation: {e}")
finally:
    excel.Quit()
