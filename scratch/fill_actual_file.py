import openpyxl
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

file_path = r"C:\Users\lap4all\Downloads\NTB_Input_Con_Thieu_Theo_Template_FLM_CRC.xlsx"

wb = openpyxl.load_workbook(file_path, data_only=False)
sheet = wb['NTB – Input']

months = ['T7', 'T8', 'T9', 'T10', 'T11', 'T12']
cols_input = ['D', 'E', 'F', 'G', 'H', 'I']
cols_db = ['I', 'J', 'K', 'L', 'M', 'N']
cols_cp = ['J', 'K', 'L', 'M', 'N', 'O']

for idx in range(6):
    ci = cols_input[idx]
    cdb = cols_db[idx]
    ccp = cols_cp[idx]
    
    # 1. Đơn GIAO – Band 3 (10–15kg) -> Row 8
    sheet[f"{ci}8"] = f"='Định biên & Sản lượng'!{cdb}72*4"
    # 2. Đơn GIAO – Band 4 (15–20kg) -> Row 9
    sheet[f"{ci}9"] = f"='Định biên & Sản lượng'!{cdb}73*4"
    
    # 3. Đơn LẤY – Band 3 -> Row 11
    sheet[f"{ci}11"] = f"='Định biên & Sản lượng'!{cdb}75*4"
    # 4. Đơn LẤY – Band 4 -> Row 12
    sheet[f"{ci}12"] = f"='Định biên & Sản lượng'!{cdb}76*4"
    # 5. Đơn LẤY – Band 5 -> Row 13
    sheet[f"{ci}13"] = f"='Định biên & Sản lượng'!{cdb}77*4"
    
    # 6. NS LẤY / 1 NV kênh nhẹ -> Row 17 (always 60)
    sheet[f"{ci}17"] = "='Định biên & Sản lượng'!$C$7"
    
    # 7. Lương NV theo xe -> Row 20 (always 0)
    sheet[f"{ci}20"] = "='Định biên & Sản lượng'!$C$34"
    
    # 8. NS NV xử lý -> Row 22 (Năng suất xử lý = 55)
    sheet[f"{ci}22"] = 55.0
    
    # 9. NS LẤY / 1 xe -> Row 27 (always 80)
    sheet[f"{ci}27"] = "='Định biên & Sản lượng'!$C$9"
    
    # 10. B2. Chi phí NV theo xe -> Row 32 (always 0)
    sheet[f"{ci}32"] = f"='Chi phí FLM'!{ccp}24*4/1000000"
    
    # 11. C2. Utilities -> Row 37 (Electricity water garbage = 5M per hub, so 5M * 4 = 20M = 20.0 million)
    sheet[f"{ci}37"] = f"='Chi phí FLM'!{ccp}33*4/1000000"
    
    # 12. D1. Setup mở mới -> Row 38 (not in data, set to 0)
    sheet[f"{ci}38"] = 0.0
    
    # 13. D2. Di dời -> Row 39 (not in data, set to 0)
    sheet[f"{ci}39"] = 0.0
    
    # 14. TỔNG CHI PHÍ chiều GIAO -> Row 41
    sheet[f"{ci}41"] = f"='Chi phí FLM'!{ccp}37*4/1000000"
    # 15. TỔNG CHI PHÍ chiều NHẬN -> Row 42
    sheet[f"{ci}42"] = f"='Chi phí FLM'!{ccp}38*4/1000000"
    
    # 16. Chi phí / đơn -> Row 44
    sheet[f"{ci}44"] = f"='Chi phí FLM'!{ccp}45"
    # 17. Tổng cost / kg -> Row 45
    sheet[f"{ci}45"] = f"='Chi phí FLM'!{ccp}46"

wb.save(file_path)
print(f"File updated and saved to: {file_path}")
wb.close()
