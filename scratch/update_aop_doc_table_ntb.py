import docx
import sys
sys.stdout.reconfigure(encoding='utf-8')

in_path = r'C:\Users\lap4all\Downloads\AOP_BCCK_Plan_new_final_v2.docx'
out_path = r'C:\Users\lap4all\Downloads\AOP_BCCK_Plan_new_final_v3.docx'

doc = docx.Document(in_path)

# Correct values based on the NTB Excel calculation (AOP_Hang_NTB_T7-T12_2026_calculated.xlsx)
correct_data = {
    'BCCK Nha Trang': ['758 đơn\n13 chuyến\n7 xe', '1.166 đơn\n20 chuyến\n10 xe', '1.316 đơn\n22 chuyến\n12 xe'],
    'BCCK Di Linh': ['644 đơn\n11 chuyến\n6 xe', '984 đơn\n17 chuyến\n9 xe', '1.114 đơn\n19 chuyến\n10 xe'],
    'BCCK Đơn Dương': ['173 đơn\n3 chuyến\n2 xe', '262 đơn\n4 chuyến\n3 xe', '297 đơn\n5 chuyến\n3 xe'],
    'BCCK Đức Linh': ['81 đơn\n1 chuyến\n1 xe', '122 đơn\n2 chuyến\n2 xe', '139 đơn\n2 chuyến\n2 xe']
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
print(f"Đã cập nhật bảng số liệu NTB mới vào file: {out_path}")
