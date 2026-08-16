import docx
import sys
sys.stdout.reconfigure(encoding='utf-8')

in_path = r'C:\Users\lap4all\Downloads\AOP_BCCK_Plan_new_final.docx'
out_path = r'C:\Users\lap4all\Downloads\AOP_BCCK_Plan_new_final_v2.docx'

doc = docx.Document(in_path)

# Correct values for T7, T10, T12
correct_data = {
    'BCCK Nha Trang': ['424 đơn\n9 chuyến\n5 xe', '849 đơn\n17 chuyến\n9 xe', '1.041 đơn\n20 chuyến\n10 xe'],
    'BCCK Di Linh': ['273 đơn\n6 chuyến\n3 xe', '344 đơn\n7 chuyến\n4 xe', '421 đơn\n9 chuyến\n5 xe'],
    'BCCK Đơn Dương': ['139 đơn\n3 chuyến\n2 xe', '183 đơn\n4 chuyến\n3 xe', '224 đơn\n5 chuyến\n3 xe'],
    'BCCK Đức Linh': ['111 đơn\n3 chuyến\n2 xe', '158 đơn\n4 chuyến\n2 xe', '194 đơn\n4 chuyến\n3 xe']
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
print(f"Đã cập nhật bảng số liệu chính xác vào file: {out_path}")
