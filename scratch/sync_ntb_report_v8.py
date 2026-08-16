import docx
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

src_path = r"C:\Users\lap4all\Downloads\AOP_Hang_Nang_NTB_T7-T12_2026_v8_8.docx"
dst_path = r"C:\Users\lap4all\Downloads\AOP_Hang_Nang_NTB_T7-T12_2026_v8_9.docx"

if not os.path.exists(src_path):
    print(f"File nguồn không tồn tại: {src_path}")
    sys.exit(1)

doc = docx.Document(src_path)

print("=== BẮT ĐẦU CẬP NHẬT ĐỢT 9 ===")

table_8 = doc.tables[7]  # Bảng 8 (index 7)

# Cập nhật các dòng của Bảng 8:
# Row 2: Tổng đầu xe BQ
bq_cars = ["20.4", "22.6", "31.5", "32.5", "37.2", "38.0"]
for c_idx in range(1, 7):
    table_8.rows[2].cells[c_idx].text = bq_cars[c_idx - 1]

# Row 3: Tổng đầu xe ngày cao điểm
peak_cars = ["30", "31", "41", "42", "51", "51"]
for c_idx in range(1, 7):
    table_8.rows[3].cells[c_idx].text = peak_cars[c_idx - 1]

# Row 4: Số người giao (đỉnh × 2 người/xe)
people_giao = ["60", "62", "82", "84", "102", "102"]
for c_idx in range(1, 7):
    table_8.rows[4].cells[c_idx].text = people_giao[c_idx - 1]

# Row 5: Mặt bằng tổng 4 BCCK cần (m²)
mb_needs = ["550", "584", "834", "850", "1,017", "1,034"]
for c_idx in range(1, 7):
    table_8.rows[5].cells[c_idx].text = mb_needs[c_idx - 1]

print("-> Đã cập nhật xong Bảng 8 theo đúng số xe bình quân, xe đỉnh và người giao từ Excel v18")

doc.save(dst_path)
print(f"=== ĐÃ LƯU FILE THÀNH CÔNG: {dst_path} ===")
