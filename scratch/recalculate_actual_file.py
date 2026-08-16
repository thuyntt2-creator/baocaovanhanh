import win32com.client
import sys

sys.stdout.reconfigure(encoding='utf-8')

file_path = r"C:\Users\lap4all\Downloads\NTB_Input_Con_Thieu_Theo_Template_FLM_CRC.xlsx"

print("Opening Excel to force recalculation of the actual file...")
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Open(file_path)
    wb.Save()
    wb.Close()
    print("Recalculation and save complete.")
except Exception as e:
    print(f"Error during Excel COM call: {e}")
finally:
    excel.Quit()
