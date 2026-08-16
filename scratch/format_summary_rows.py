import openpyxl
import win32com.client
import sys
import os

sys.stdout.reconfigure(encoding='utf-8')

downloads_dir = r"C:\Users\lap4all\Downloads"
files_to_format = [
    os.path.join(downloads_dir, "NTB_Input_Da_Dien_FLM_CRC_2.xlsx"),
    os.path.join(downloads_dir, "NTB_Input_Con_Thieu_Theo_Template_FLM_CRC.xlsx")
]

cols = ['D', 'E', 'F', 'G', 'H', 'I']

for filepath in files_to_format:
    if not os.path.exists(filepath):
        continue
    
    print(f"Formatting cells in: {os.path.basename(filepath)}")
    wb = openpyxl.load_workbook(filepath, data_only=False)
    sheet = wb['NTB – Input']
    
    # Format Row 44 (Chi phí / đơn) -> Number format #,##0 (Integer)
    # Format Row 45 (Tổng cost / kg) -> Number format #,##0.00 (Two decimal places)
    for col in cols:
        sheet[f"{col}44"].number_format = '#,##0'
        sheet[f"{col}45"].number_format = '#,##0.00'
        
    wb.save(filepath)
    wb.close()
    
    # Force Excel COM to recalculate and save formatting
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        workbook = excel.Workbooks.Open(filepath)
        workbook.Save()
        workbook.Close()
        print("Success!")
    except Exception as e:
        print(f"Error: {e}")
    finally:
        excel.Quit()
