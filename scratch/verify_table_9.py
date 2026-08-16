import docx
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

doc_path = r"C:\Users\lap4all\Downloads\AOP_Hang_Nang_NTB_T7-T12_2026_v8_3.docx"

if not os.path.exists(doc_path):
    print("File không tồn tại")
    sys.exit(1)

doc = docx.Document(doc_path)
table = doc.tables[8]  # index 8 (Bảng 9)

print("=== VERIFY BẢNG 9 ===")
errors = 0

# Kiểm tra tổng m2 cần
tot_m2 = int(table.rows[5].cells[3].text.replace(',', ''))
sum_m2 = int(table.rows[1].cells[3].text) + int(table.rows[2].cells[3].text) + int(table.rows[3].cells[3].text) + int(table.rows[4].cells[3].text)
if tot_m2 != sum_m2:
    print(f"  [LỖI] Tổng m2 cần = {tot_m2} | Thực tế sum = {sum_m2}")
    errors += 1
else:
    print(f"  [OK] Tổng m2 cần = {tot_m2} (Khớp)")

# Kiểm tra tổng giá thuê
tot_price = int(table.rows[5].cells[4].text.replace('.', '').replace(' đ/th', '').replace(' ', ''))
sum_price = int(table.rows[1].cells[4].text.replace('.', '').replace(' đ/th', '')) + \
            int(table.rows[2].cells[4].text.replace('.', '').replace(' đ/th', '')) + \
            int(table.rows[3].cells[4].text.replace('.', '').replace(' đ/th', '')) + \
            int(table.rows[4].cells[4].text.replace('.', '').replace(' đ/th', ''))
if tot_price != sum_price:
    print(f"  [LỖI] Tổng giá thuê = {tot_price} | Thực tế sum = {sum_price}")
    errors += 1
else:
    print(f"  [OK] Tổng giá thuê = {tot_price} (Khớp)")

# Kiểm tra tổng m2 thiếu
tot_lack = int(table.rows[5].cells[6].text.replace(',', ''))
sum_lack = int(table.rows[1].cells[6].text) + int(table.rows[2].cells[6].text) + int(table.rows[3].cells[6].text) + int(table.rows[4].cells[6].text)
if tot_lack != sum_lack:
    print(f"  [LỖI] Tổng m2 thiếu = {tot_lack} | Thực tế sum = {sum_lack}")
    errors += 1
else:
    print(f"  [OK] Tổng m2 thiếu = {tot_lack} (Khớp)")

if errors == 0:
    print(">>> TẤT CẢ KIỂM TRA BẢNG 9 ĐỀU THÀNH CÔNG! <<<")
else:
    print(f">>> CÓ {errors} LỖI BẢNG 9 <<<")
