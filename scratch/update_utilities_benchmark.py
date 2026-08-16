import openpyxl
import win32com.client
import sys
import os

sys.stdout.reconfigure(encoding='utf-8')

downloads_dir = r"C:\Users\lap4all\Downloads"
files_to_fix = [
    os.path.join(downloads_dir, "NTB_Input_Da_Dien_FLM_CRC_2.xlsx"),
    os.path.join(downloads_dir, "NTB_Input_Con_Thieu_Theo_Template_FLM_CRC.xlsx")
]

for filepath in files_to_fix:
    if not os.path.exists(filepath):
        continue
        
    print(f"\nUpdating Utilities benchmark in: {os.path.basename(filepath)}")
    wb = openpyxl.load_workbook(filepath, data_only=False)
    
    # 1. Update Định biên & Sản lượng cell C41 (Điện nước rác khoán) to 1,122,389.45 VNĐ
    sheet_db = wb['Định biên & Sản lượng']
    sheet_db['C41'] = 1122389.45
    
    wb.save(filepath)
    wb.close()
    
    # Recalculate via Excel COM
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

# Verify values
for filepath in files_to_fix:
    if not os.path.exists(filepath):
        continue
    wb = openpyxl.load_workbook(filepath, data_only=True)
    sheet = wb['NTB – Input']
    cols = ['D', 'E', 'F', 'G', 'H', 'I']
    print(f"\n=== VERIFIED VALUES IN: {os.path.basename(filepath)} ===")
    print("Row 37 (C2. Utilities):", [sheet[f"{c}37"].value for c in cols])
    print("Row 41 (Giao):         ", [sheet[f"{c}41"].value for c in cols])
    print("Row 42 (Nhận):         ", [sheet[f"{c}42"].value for c in cols])
    print("Row 43 (Total FLM):    ", [sheet[f"{c}43"].value for c in cols])
    print("Row 44 (Chi phí/đơn):  ", [sheet[f"{c}44"].value for c in cols])
    wb.close()
