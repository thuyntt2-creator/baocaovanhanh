import docx
import sys
sys.stdout.reconfigure(encoding='utf-8')

in_path = r'C:\Users\lap4all\Downloads\AOP_BCCK_Plan_new_final.docx'
out_path = r'C:\Users\lap4all\Downloads\AOP_BCCK_Plan_new_final_v4.docx'

doc = docx.Document(in_path)

# Correct values matching Excel file V16 (AOP_Hang_Nang_MAU_Vung_T7-T12_2026_final_v16.xlsx)
# format: [T7, T10, T12] -> 'SL đơn\nChuyến xe\nSố xe'
# We do not divide the number of cars by 2, so Số xe = Chuyến xe to match the Excel sheet's calculations exactly.
correct_data = {
    'BCCK Nha Trang': ['424 đơn\n9 chuyến\n9 xe', '849 đơn\n17 chuyến\n17 xe', '1.041 đơn\n20 chuyến\n20 xe'],
    'BCCK Di Linh': ['273 đơn\n6 chuyến\n6 xe', '344 đơn\n7 chuyến\n7 xe', '421 đơn\n9 chuyến\n9 xe'],
    'BCCK Đơn Dương': ['139 đơn\n3 chuyến\n3 xe', '183 đơn\n4 chuyến\n4 xe', '224 đơn\n5 chuyến\n5 xe'],
    'BCCK Đức Linh': ['111 đơn\n3 chuyến\n3 xe', '158 đơn\n4 chuyến\n4 xe', '194 đơn\n4 chuyến\n4 xe']
}

for table in doc.tables:
    for row in table.rows:
        cell_0_text = row.cells[0].text.strip()
        if cell_0_text in correct_data and len(row.cells) >= 5:
            vals = correct_data[cell_0_text]
            row.cells[2].text = vals[0]
            row.cells[3].text = vals[1]
            row.cells[4].text = vals[2]

doc.save(out_path)
print(f"Đã cập nhật số liệu chuẩn từ Excel V16 vào file: {out_path}")
