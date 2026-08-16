import win32com.client
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

file_name = "V2 AOP_Hang_NTB_T7-T12_2026.xlsx"
file_path = os.path.join(r"C:\Users\lap4all\Downloads", file_name)

if not os.path.exists(file_path):
    print(f"Error: {file_path} does not exist!")
    sys.exit(1)

print("Connecting to Excel application...")
try:
    # Try getting the active Excel instance or starting a new one
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
        print("Connected to active Excel instance.")
    except Exception:
        excel = win32com.client.Dispatch("Excel.Application")
        print("Started new Excel application instance.")
        
    excel.Visible = True
    excel.DisplayAlerts = False
    
    print(f"Opening workbook: {file_path}...")
    wb = excel.Workbooks.Open(file_path)
    
    print("Forcing full recalculation...")
    excel.CalculateFull()
    
    print("Saving workbook...")
    wb.Save()
    print("✅ Workbook calculated and saved successfully!")
    
except Exception as e:
    print(f"\n❌ Error during Excel calculation: {e}")
    sys.exit(1)
