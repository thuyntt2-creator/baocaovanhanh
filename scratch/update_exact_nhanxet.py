import os, sys

workspace_md_path = r'c:\Users\lap4all\Documents\Auto report\KTC_Ke_Hoach_Van_Hanh_Event_8_8_NTB.md'
artifact_md_path = r'C:\Users\lap4all\.gemini\antigravity-ide\brain\497c8661-37a7-4c4e-986e-e8ef6c144fb3\KTC_Ke_Hoach_Van_Hanh_Event_8_8_NTB.md'

nhan_xet_section = """
### 2.3. Nhận Xét Chuyên Sâu Chi Tiết Cho Từng Kho (Chuẩn hóa 100% Theo Mẫu TNB)

![Cơ cấu Nhóm hàng tại KTC NTB](file:///C:/Users/lap4all/.gemini/antigravity-ide/brain/497c8661-37a7-4c4e-986e-e8ef6c144fb3/chart2_product_group_breakdown.png)

#### **- Nhận xét:**

##### **KTC Khánh Hòa:**
- **• XU HƯỚNG**: Biến động sóng kép qua kỳ theo dõi: TB **29,373 đơn/ngày**. Đỉnh campaign 08/08 đạt **36,898 đơn** (+25.6% so với TB kỳ) do Bulky vọt lên 3,940 đơn và Freight 2,901 đơn (tổng Bulky+Freight đạt 6,841 đơn, gấp 1.2 lần nền ngày thường ~5.9K); Đỉnh 2 (10/08) đạt **33,041 đơn** (+12.5%); Đáy Chủ Nhật 09/08 đạt 25,208 đơn (−14.2%). Biên độ đỉnh/đáy **1.46 lần**.
- **• CƠ CẤU**: Normal **79.7% (234,077 đơn)** – Bulky **12.6% (37,008 đơn)** – Freight **7.7% (22,646 đơn)**; Bulky của Khánh Hòa chiếm khối lượng tuyệt đối **CAO NHẤT VÙNG (37,008 đơn/10 ngày)** — gấp 1.9 lần tổng Bulky của Đức Trọng và gấp 2.8 lần Bảo Lộc.
- **• DỰ KIẾN & CHUẨN BỊ**: Ngày thường 26–30K, đầu tuần (T2–T3) 29–33K, Chủ Nhật ~25K; nếu lặp kịch bản campaign, đỉnh có thể chạm **38–40K/ngày** với **>7K đơn Bulky+Freight** — cần dự phòng +30% nhân lực phân loại, quy hoạch diện tích sàn hàng cồng kềnh/hàng nặng và bố trí đầu xe tải trọng lớn (15–20 tấn) bọc lót theo **NGÀY ĐỈNH (~36.9K đơn/ngày)**, không theo trung bình (~29.4K/ngày).

##### **CT Bình Thuận:**
- **• XU HƯỚNG**: TB **17,485 đơn/ngày**. Đỉnh campaign 08/08 đạt **21,946 đơn** (+25.5% so với TB kỳ) do Bulky vọt lên 2,347 đơn và Freight 2,334 đơn (tổng Bulky+Freight 4,681 đơn, chiếm 21.3% sản lượng ngày); Đỉnh 2 (10/08) đạt **20,302 đơn** (+16.1%); Đáy Chủ Nhật 09/08 đạt 14,541 đơn (−16.8%). Biên độ đỉnh/đáy **1.51 lần**.
- **• CƠ CẤU**: Normal **76.9%** – Bulky **13.6%** – Freight **9.5%**; tỷ trọng Freight (nông sản/hải sản) **CAO NHẤT VÙNG (16,618 đơn/10 ngày)** — phản ánh đặc thù cửa ngõ phía Nam kết nối trực tiếp các tuyến xe liên vùng từ TP.HCM/Đồng Nai.
- **• DỰ KIẾN & CHUẨN BỊ**: Ngày thường 16–18K, Chủ Nhật ~14.5K; đợt campaign có thể chạm **22–24K/ngày** với **~5K đơn Bulky+Freight** dồn trong ngày đỉnh — đặc biệt lưu ý định mức năng suất nhân sự rất cao (11 NVCT), cần chuẩn bị trước 3-4 Freelance ca đêm và quy hoạch luồng xả bãi pallet nhanh cho xe liên vùng về trễ.

##### **CT Đức Trọng (Lâm Đồng):**
- **• XU HƯỚNG**: TB **14,700 đơn/ngày**. Đỉnh campaign 08/08 đạt **20,017 đơn** (+36.2% so với TB kỳ - **biên độ tăng cao nhất vùng**) do Bulky tăng vọt lên **3,483 đơn** (gấp 2.4 lần nền ngày thường ~1.4K) và Freight 2,032 đơn; Đỉnh 2 (10/08) đạt **17,743 đơn** (+20.7%); Đáy Chủ Nhật 09/08 đạt 12,172 đơn (−17.2%). Biên độ đỉnh/đáy **1.64 lần (cao nhất NTB)**.
- **• CƠ CẤU**: Normal **76.8%** – Bulky **13.3%** – Freight **9.8%**; tổng hàng cồng kềnh & nặng đạt **34,057 đơn/10 ngày** — đóng vai trò Hub trung chuyển chính cho toàn tỉnh Lâm Đồng (Đà Lạt & Bảo Lộc).
- **• DỰ KIẾN & CHUẨN BỊ**: Ngày thường 13–15K, Chủ Nhật ~12K; đợt campaign có thể chạm **21–22K/ngày** với **>5.5K đơn Bulky+Freight** — cần bố trí sẵn bãi hạ pallet cồng kềnh riêng biệt tại cửa xả hàng, bổ sung 2 xe nâng điện 2 tầng và điều phối tuyến xe tải lớn kết nối Đức Trọng - Đà Lạt.

##### **CT Bảo Lộc (Lâm Đồng):**
- **• XU HƯỚNG**: TB **9,565 đơn/ngày**. Đỉnh campaign 08/08 đạt **11,254 đơn** (+17.7% so với TB kỳ) với Bulky 1,368 đơn và Freight 1,112 đơn; Đỉnh 2 (10/08) đạt **11,541 đơn** (+20.7%); Đáy Chủ Nhật 09/08 đạt 8,423 đơn (−11.9%). Biên độ đỉnh/đáy **1.37 lần**.
- **• CƠ CẤU**: Normal **77.5%** – Bulky **13.6% (12,993 đơn)** – Freight **8.9% (8,524 đơn)**; cơ cấu khá ổn định qua các ngày trong đợt.
- **• DỰ KIẾN & CHUẨN BỊ**: Ngày thường 9–10K, Chủ Nhật ~8.4K; đợt campaign chạm **11.5–12K/ngày** — ưu tiên cơ chế ca kíp linh hoạt (dùng 7 Freelance/ngày) và gom chuyến điều phối xe chung tuyến với Đức Trọng thay vì tăng nguồn lực cố định.

##### **CT Đắc Nông:**
- **• XU HƯỚNG**: TB **6,075 đơn/ngày**. Đỉnh campaign 08/08 đạt **7,534 đơn** (+24.0% so với TB kỳ) với Bulky 887 đơn và Freight 529 đơn; Đỉnh 2 (10/08) đạt **6,981 đơn** (+14.9%); Đáy Chủ Nhật 09/08 đạt 5,216 đơn (−14.1%). Biên độ đỉnh/đáy **1.44 lần**.
- **• CƠ CẤU**: Normal **78.4%** – Bulky **14.1% (8,592 đơn - tỷ trọng Bulky cao nhất vùng)** – Freight **7.4% (4,513 đơn)**.
- **• DỰ KIẾN & CHUẨN BỊ**: Ngày thường 5.5–6.2K, Chủ Nhật ~5.2K; đợt campaign chạm **7.5–8K/ngày** — quy mô nhỏ nhưng hàng Bulky chiếm tỷ trọng cao; duy trì lực lượng 6 Freelance bọc lót và kiểm tra 100% PDA + SIM 4G dự phòng trước ca cao điểm.
"""

# Read existing files and replace section 2.3
with open(workspace_md_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Replace section 2.3
import re
new_content = re.sub(
    r'### 2\.3\. Nhận Xét Chuyên Sâu.*?(?=### 2\.4\.)',
    nhan_xet_section + '\n\n',
    content,
    flags=re.DOTALL
)

with open(workspace_md_path, 'w', encoding='utf-8') as f:
    f.write(new_content)

with open(artifact_md_path, 'w', encoding='utf-8') as f:
    f.write(new_content)

print('Updated Markdown files with exact 3-bullet TNB Nhận xét structure!')
