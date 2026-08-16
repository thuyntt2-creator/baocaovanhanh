import docx
import os
import sys
sys.stdout.reconfigure(encoding='utf-8')

in_path = r'C:\Users\lap4all\Downloads\AOP_BCCK_Plan_new.docx'
out_path = r'C:\Users\lap4all\Downloads\AOP_BCCK_Plan_new_final.docx'

doc = docx.Document(in_path)

# Prepare new text blocks
new_6_1 = """Nhân sự được định biên linh hoạt theo sản lượng dự báo (Forecast) thực tế từng tháng của từng trạm, không cào bằng định mức tĩnh như phương án cũ.
Công thức tính nhân sự cho từng BCCK:
- Nhân sự 1 BCCK = 3 Nhân viên Xử lý (cố định) + ROUNDUP(Sản lượng dự báo ngày / Năng suất định mức 1 shipper).
- Hàng vừa (10–20kg) được phân rã về bưu cục thường chặng cuối cho shipper thường (xe máy), giúp giải phóng áp lực cho hệ thống BCCK."""

new_7_1 = """Để đảm bảo ngân sách Vùng phản ánh chính xác nguồn lực thực tế của 4 BCCK, Cấu trúc Chi phí Topline (Dòng 17 - Chi phí nhân sự xử lý) đã được điều chỉnh:
- KHÔNG sử dụng hàm tính nhân sự xử lý mặc định của công ty (dựa trên tổng volume vùng) vì không sát thực tế 4 Hub.
- THAY THẾ bằng Tổng chi phí vận hành trực tiếp của 4 Hub, bao gồm:
  + Chi phí Lương: Nhân sự từng BCCK × 15.000.000 đ/người/tháng.
  + Chi phí Xe tải 1.9T: Xe BQ/ngày của từng BCCK × 1.200.000 đ/xe × số ngày làm việc.
- Dòng 17 trên Topline = Tổng (Lương + Xe) của cả 4 BCCK. Việc gom nhóm này giúp dễ dàng theo dõi hiệu quả tài chính độc lập của dự án BCCK trên báo cáo tổng thể mà không làm lặp (double-count) chi phí chung của công ty.
- Chi phí mặt bằng (Dòng 15) và Chi phí xe chung (Dòng 16): Giữ nguyên cấu trúc phân bổ của Vùng để chi trả cho các hoạt động vận tải ngoài phạm vi 4 Hub."""

new_7_2 = """Tổng chi phí tính toán sát với sản lượng T7-T12/2026:
- T7: 2,109,900,000 VNĐ
- T8: 2,293,800,000 VNĐ
- T9: 3,008,100,000 VNĐ
- T10: 3,138,300,000 VNĐ
- T11: 3,521,550,000 VNĐ
- T12: 3,702,300,000 VNĐ
* Chi tiết tính toán tự động ăn theo công thức được trình bày ở sheet "Nguồn lực & chi phí" trong file AOP_Hang_Nang_MAU_Vung_T7-T12_2026_final_v20.xlsx."""

for para in doc.paragraphs:
    if "Nhân sự được định biên theo sản lượng thực tế Topline H2" in para.text:
        para.text = new_6_1
    if "Chi phí xe tải 1.9T: xe BQ/ngày" in para.text:
        para.text = new_7_1
    if "Chi phí nhân sự giao hàng (38 shipper): lương + phụ cấp" in para.text:
        para.text = ""
    if "Chi phí nhân sự kho & quản lý (8 người cố định" in para.text:
        para.text = ""
    if "Chi phí thuê mặt bằng: diện tích × 120.000" in para.text:
        para.text = ""
    if "* Số liệu ước tính. Chi phí chính xác cập nhật" in para.text:
        para.text = new_7_2

doc.save(out_path)
print(f"Đã lưu file báo cáo hoàn chỉnh tại: {out_path}")
