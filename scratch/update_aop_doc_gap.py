import docx
import sys
sys.stdout.reconfigure(encoding='utf-8')

in_path = r'C:\Users\lap4all\Downloads\AOP_BCCK_Plan_new_final.docx'
out_path = r'C:\Users\lap4all\Downloads\AOP_BCCK_Plan_new_final.docx'

doc = docx.Document(in_path)

# Prepare the new section text
gap_text = """5.5 Logic xác định Ngày thường & Ngày cao điểm (Peak) - Chứng minh Năng lực GAP
- Sản lượng bình quân/ngày: Được tính toán sát thực tế bằng hàm trung bình (AVERAGE) của 30-31 ngày từ dữ liệu Forecast chi tiết từng ngày.
- Sản lượng đỉnh/ngày (Peak): Được xác định bằng ngày có sản lượng đổ về cao nhất (MAX) trong tháng.
- Chứng minh Năng lực GAP: Hiện tại, năng lực của hệ thống bưu cục tuyến nội thị chỉ đáp ứng được sản lượng bình quân. Khi rơi vào ngày Peak hoặc các dịp Mega Sales, sản lượng tăng đột biến tạo ra GAP âm từ 500 – 1.500 đơn/ngày. Do hàng nặng (Bulky) chiếm diện tích lớn và cần xe tải, nếu để tại bưu cục cũ sẽ gây "nghẽn cổ chai" toàn hệ thống. Việc quy hoạch 4 BCCK riêng biệt với định biên nhân sự linh hoạt (cộng thêm NVXL và xe tải) giúp hấp thụ hoàn toàn phần GAP này, đảm bảo vận hành trơn tru cả trong ngày thường lẫn ngày đỉnh, triệt tiêu hoàn toàn rủi ro quá tải."""

# We will insert this at the end of Section V (before Section VI)
for i, para in enumerate(doc.paragraphs):
    if "VI. ĐÌNH BIÊN NHÂN SỰ" in para.text:
        # Insert before this paragraph
        p = para.insert_paragraph_before(gap_text)
        p.style = doc.styles['Normal']
        # Add an empty paragraph for spacing
        para.insert_paragraph_before("")
        break

doc.save(out_path)
print(f"Đã cập nhật logic GAP vào file: {out_path}")
