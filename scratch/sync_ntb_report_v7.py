import docx
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

src_path = r"C:\Users\lap4all\Downloads\AOP_Hang_Nang_NTB_T7-T12_2026_v8_7.docx"
dst_path = r"C:\Users\lap4all\Downloads\AOP_Hang_Nang_NTB_T7-T12_2026_v8_8.docx"

if not os.path.exists(src_path):
    print(f"File nguồn không tồn tại: {src_path}")
    sys.exit(1)

doc = docx.Document(src_path)

print("=== BẮT ĐẦU CẬP NHẬT ĐỢT 8 ===")

table_3 = doc.tables[2]  # Bảng 3

# 1. Nha Trang
# Row 2: Chuyến xe. Col 2: T7, Col 3: T10, Col 4: T12
table_3.rows[2].cells[2].text = "28 chuyến"
table_3.rows[2].cells[3].text = "46 chuyến"
table_3.rows[2].cells[4].text = "56 chuyến"
# Row 3: Số xe
table_3.rows[3].cells[2].text = "14 xe"
table_3.rows[3].cells[3].text = "23 xe"
table_3.rows[3].cells[4].text = "28 xe"

# 2. Di Linh
# Row 5: Chuyến xe
table_3.rows[5].cells[2].text = "16 chuyến"
table_3.rows[5].cells[3].text = "20 chuyến"
table_3.rows[5].cells[4].text = "24 chuyến"
# Row 6: Số xe
table_3.rows[6].cells[2].text = "8 xe"
table_3.rows[6].cells[3].text = "10 xe"
table_3.rows[6].cells[4].text = "12 xe"

# 3. Đơn Dương
# Row 8: Chuyến xe
table_3.rows[8].cells[2].text = "8 chuyến"
table_3.rows[8].cells[3].text = "12 chuyến"
table_3.rows[8].cells[4].text = "14 chuyến"
# Row 9: Số xe
table_3.rows[9].cells[2].text = "4 xe"
table_3.rows[9].cells[3].text = "6 xe"
table_3.rows[9].cells[4].text = "7 xe"

# 4. Đức Linh
# Row 11: Chuyến xe
table_3.rows[11].cells[2].text = "8 chuyến"
table_3.rows[11].cells[3].text = "10 chuyến"
table_3.rows[11].cells[4].text = "12 chuyến"
# Row 12: Số xe
table_3.rows[12].cells[2].text = "4 xe"
table_3.rows[12].cells[3].text = "5 xe"
table_3.rows[12].cells[4].text = "6 xe"

print("-> Đã cập nhật xong Bảng 3 theo đúng xe đỉnh của plan Excel v18")

doc.save(dst_path)
print(f"=== ĐÃ LƯU FILE THÀNH CÔNG: {dst_path} ===")
