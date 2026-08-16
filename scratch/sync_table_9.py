import docx
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

doc_path = r"C:\Users\lap4all\Downloads\AOP_Hang_Nang_NTB_T7-T12_2026_v8_3.docx"

if not os.path.exists(doc_path):
    print(f"File không tồn tại: {doc_path}")
    sys.exit(1)

doc = docx.Document(doc_path)
table = doc.tables[8]  # index 8 (Bảng 9)

print("=== BẮT ĐẦU CẬP NHẬT BẢNG 9 ===")

# Sửa Row 1: Nha Trang
table.rows[1].cells[4].text = "41.760.000 đ/th"
table.rows[1].cells[5].text = ""
table.rows[1].cells[6].text = "348"

# Sửa Row 2: Di Linh
table.rows[2].cells[4].text = "17.040.000 đ/th"
table.rows[2].cells[5].text = ""
table.rows[2].cells[6].text = "142"

# Sửa Row 3: Đơn Dương
table.rows[3].cells[4].text = "9.000.000 đ/th"
table.rows[3].cells[5].text = ""
table.rows[3].cells[6].text = "75"

# Sửa Row 4: Đức Linh
table.rows[4].cells[4].text = "7.800.000 đ/th"
table.rows[4].cells[5].text = ""
table.rows[4].cells[6].text = "65"

# Sửa Row 5: TỔNG
table.rows[5].cells[4].text = "75.600.000 đ/th"
table.rows[5].cells[5].text = ""
table.rows[5].cells[6].text = "630"

print("-> Đã cập nhật xong các ô trong Bảng 9")

doc.save(doc_path)
print(f"=== ĐÃ LƯU FILE THÀNH CÔNG: {doc_path} ===")
