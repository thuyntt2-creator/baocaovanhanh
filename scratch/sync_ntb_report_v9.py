import docx
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

src_path = r"C:\Users\lap4all\Downloads\AOP_Hang_Nang_NTB_T7-T12_2026_v8_9.docx"
dst_path = r"C:\Users\lap4all\Downloads\AOP_Hang_Nang_NTB_T7-T12_2026_v8_10.docx"

if not os.path.exists(src_path):
    print(f"File nguồn không tồn tại: {src_path}")
    sys.exit(1)

doc = docx.Document(src_path)

print("=== BẮT ĐẦU CẬP NHẬT ĐỢT 10 ===")

table_8 = doc.tables[7]  # Bảng 8

# Cập nhật dòng 2: Tổng đầu xe BQ (làm tròn số nguyên)
bq_cars = ["21", "23", "31", "32", "37", "38"]
for c_idx in range(1, 7):
    table_8.rows[2].cells[c_idx].text = bq_cars[c_idx - 1]

print("-> Đã sửa xe bình quân thành số nguyên tròn ở Bảng 8")

doc.save(dst_path)
print(f"=== ĐÃ LƯU FILE THÀNH CÔNG: {dst_path} ===")
