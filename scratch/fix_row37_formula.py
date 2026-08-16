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

cols = ['D', 'E', 'F', 'G', 'H', 'I']
flm_cols = ['J', 'K', 'L', 'M', 'N', 'O']

for filepath in files_to_fix:
    if not os.path.exists(filepath):
        continue
        
    print(f"Applying dynamic rounded formula to Row 37 in: {os.path.basename(filepath)}")
    wb = openpyxl.load_workbook(filepath, data_only=False)
    sheet = wb['NTB – Input']
    
    # Update Row 37 (C2. Utilities)
    for idx, col in enumerate(cols):
        f_col = flm_cols[idx]
        sheet[f"{col}37"] = f"=ROUND('Chi phí FLM'!{f_col}33*4/1000000, 2)"
        
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
        print(f"Error during recalculation: {e}")
    finally:
        excel.Quit()
        
# Verify the rounded values
for filepath in files_to_fix:
    if not os.path.exists(filepath):
        continue
    wb = openpyxl.load_workbook(filepath, data_only=True)
    sheet = wb['NTB – Input']
    print(f"\n=== VERIFIED VALUES IN: {os.path.basename(filepath)} ===")
    print("Row 37 (C2. Utilities):", [sheet[f"{c}37"].value for c in cols])
    print("Row 43 (Total FLM):    ", [sheet[f"{c}43"].value for c in cols])
    wb.close()
