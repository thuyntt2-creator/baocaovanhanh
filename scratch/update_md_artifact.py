import pandas as pd, sys, os

sys.stdout.reconfigure(encoding='utf-8')

artifact_dir = r"C:\Users\lap4all\.gemini\antigravity-ide\brain\2488f1f5-238f-4151-b3dd-1b59691a739c"
md_path = os.path.join(artifact_dir, "bao_cao_quy_hoach_ntb_dvhc_moi.md")

df_excel = pd.read_excel(r'C:\Users\lap4all\Downloads\NTB_Phan_Tuyen_Hanh_Chinh_Quy_Hoach_Moi.xlsx', sheet_name='Sheet1')

md = []
md.append("# BÁO CÁO TOÀN DIỆN & PHƯƠNG ÁN QUY HOẠCH MẠNG LƯỚI BƯU CỤC VÙNG NAM TRUNG BỘ (NTB) THEO ĐƠN VỊ HÀNH CHÍNH MỚI")
md.append("\n*Đồng bộ Số liệu Sản lượng AM & Phân tích Chi tiết Lý do Giữ nguyên 2-3 Bưu cục Không gộp (Kèm Bản đồ Quy hoạch 2026)*\n")

md.append("> [!NOTE]")
md.append("> File báo cáo Word (.docx) hoàn chỉnh v3 đã đồng bộ toàn bộ sản lượng AM và Lý do Giữ nguyên 2-3 Bưu cục không gộp đã được xuất tại:")
md.append("> `C:\\Users\\lap4all\\Downloads\\Bao_Cao_Quy_Hoach_Buu_Cuc_NTB_Theo_DVHC_Moi_V3.docx`\n")

md.append("---")
md.append("\n## I. TỔNG QUAN HIỆN TRẠNG MẠNG LƯỚI BƯU CỤC VÙNG NTB\n")
md.append("Thực hiện chủ trương tinh gọn và sáp nhập đơn vị hành chính cấp xã/phường trên phạm vi toàn quốc năm 2026, mạng lưới giao nhận bưu cục vùng Nam Trung Bộ (bao gồm 5 tỉnh: **Khánh Hòa, Ninh Thuận, Bình Thuận, Lâm Đồng, Đắc Nông**) đứng trước yêu cầu tái cấu trúc toàn diện phạm vi quản lý tuyến, diện tích kho bãi và bố trí nhân sự.\n")

md.append("### 1. Quy mô Mạng lưới & Mức độ Chia cắt Địa bàn")
md.append("- **Tổng số Bưu cục Express đang vận hành:** `83 Bưu cục` trên địa bàn 5 tỉnh.")
md.append("- **Phạm vi rà soát hành chính mới:** `36 Xã/Phường mới` (được sáp nhập từ **114 xã/phường cũ**), đồng bộ chuẩn 100% theo file `NTB_Phan_Tuyen_Hanh_Chinh_Quy_Hoach_Moi.xlsx`.")
md.append("- **Số Xã/Phường đã quy hoạch chuẩn (01 BC duy nhất phụ trách):** `4 Xã/Phường` (đạt **11.1%**) như Xã Cam Hiệp, Xã Cam Lâm, Xã Phước Dinh, Xã Phan Rí Cửa.")
md.append("- **Số Xã/Phường bị CHIA CẮT MẢNH (có 2 - 3 BC cùng giao hàng):** `32 Xã/Phường` (chiếm **88.9%**).\n")

md.append("### 2. Các Bất cập & Rủi ro Vận hành Chính\n")
md.append("> [!WARNING]")
md.append("> **3 Thách thức lớn khi sáp nhập hành chính mới:**")
md.append("> 1. **Chồng chéo tuyến đường & dẫm chân shipper:** Tại 32 phường/xã mới bị chia cắt, có từ 2 đến 3 Bưu cục cùng cử shipper vào cùng một phường mới để giao hàng làm gia tăng chi phí xăng xe và thời gian di chuyển.")
md.append("> 2. **Tuyến đi chéo xa ranh giới:** Xuất hiện hiện tượng một số xã vùng ven bị phân tuyến giao từ Bưu cục ở huyện/tỉnh khác cách xa **30 - 40 km**, trong khi Bưu cục lân cận chỉ cách **7 - 15 km**.")
md.append("> 3. **Rủi ro đứt gãy nhân sự tại địa bàn miền núi/nông thôn:** Nếu cưỡng ép đóng cửa hoặc di dời BC sáp nhập ngay lập tức theo địa giới mới (ví dụ tại Đắk Nông), **100% nhân sự điểm xã ven khẳng định sẽ nghỉ việc** do không chấp nhận di chuyển xa.\n")

md.append("---")
md.append("\n## II. DANH MỤC 36 XÃ/PHƯỜNG MỚI THEO CHUẨN MASTER DATASET\n")
md.append("Toàn bộ tên Phường/Xã mới, Mã xã mới (GHN Code) và Tỉnh/thành mới dưới đây được đồng bộ chính xác 100% theo file danh mục chuẩn `NTB_Phan_Tuyen_Hanh_Chinh_Quy_Hoach_Moi.xlsx`:\n")

md.append("| STT | Mã Xã Mới | Tên Xã/Phường Mới (Chuẩn Excel) | Tỉnh/Thành Mới | Sản Lượng AM (Đơn/ngày) | AM Phụ Trách & Phương Án Đề Xuất |")
md.append("| :--- | :--- | :--- | :--- | :--- | :--- |")

grouped_master = df_excel.groupby(['Mã Xã mới', 'Tên Xã mới', 'Tỉnh, thành phố mới'], sort=False)
stt = 1
for (code, name, prov), grp in grouped_master:
    vol_sum = grp['Sản lượng giao/ngày (đơn)'].sum() + grp['Sản lượng lấy/ngày (đơn)'].sum()
    am = grp['Quản lý khu vực (AM)'].iloc[0]
    proposal = str(grp['Đánh giá & Phương án đề xuất'].iloc[0])
    md.append(f"| {stt} | `{code}` | **{name}** | {prov} | {vol_sum:.0f} | AM {am}: {proposal} |")
    stt += 1

md.append("\n---")
md.append("\n## III. PHƯƠNG ÁN QUY HOẠCH CHI TIẾT & LÝ DO GIỮ NGUYÊN KHÔNG GỘP CỦA AM (KÈM BẢN ĐỒ AM)\n")

# Regional content with exact AM numbers and reasons
md.append("### 1. Tỉnh Lâm Đồng\n")
md.append("#### 1.1. Thành phố Đà Lạt")
md.append("- **Phường Xuân Hương - Đà Lạt (Mã: 24781):**")
md.append("  - **Sản lượng AM trình bày:** `2.292.0 đơn/ngày` (Giao: 1.497.0 đơn · Lấy: 795.0 đơn). Hiện do 3 BC phụ trách: (LDO) Lang Biang - Đà Lạt 2, (LDO) Xuân Hương - Đà Lạt, (LDO) Lâm Viên - Đà Lạt 1.")
md.append("  - > [!IMPORTANT]")
md.append("    > **LÝ DO GIỮ NGUYÊN 2 BC & MỞ MỚI 1 BC (AM Lê Văn Trường):** Sản lượng quá lớn (2.292 đơn/ngày). Một bưu cục đơn lẻ không thể tải nổi kho bãi và nhân sự. AM đề xuất mở mới 01 BC `(LDO) Xuân Hương - Đà Lạt 2` tại Phường 10. Phân ranh giới: BC Xuân Hương 1 phụ trách Phường 1 + 2 cũ; BC Xuân Hương 2 phụ trách Phường 3 + 10 cũ.")
md.append("- **Phường Lâm Viên - Đà Lạt (Mã: 24778):**")
md.append("  - **Sản lượng AM trình bày:** `1.365.0 đơn/ngày` (Giao: 1.066.0 đơn · Lấy: 299.0 đơn). Hiện do 2 BC phụ trách: (LDO) Lâm Viên - Đà Lạt 1 và (LDO) Lâm Viên - Đà Lạt 2.")
md.append("  - > [!IMPORTANT]")
md.append("    > **LÝ DO GIỮ NGUYÊN 02 BƯU CỤC KHÔNG GỘP (AM Lê Văn Trường):** Phường Lâm Viên mới có địa hình tự nhiên bị chia cắt làm 2 bờ riêng biệt bởi Hồ Xuân Hương. Nếu gộp về 1 Bưu cục, shipper sẽ phải di chuyển vòng qua hồ rất xa để sang phía bờ đối diện, làm gia tăng thời gian di chuyển và tăng chi phí vận hành. Do đó, AM quyết định **GIỮ NGUYÊN 02 Bưu cục** (Lâm Viên 1 & Lâm Viên 2) hoạt động ở 2 phía bờ hồ để đảm bảo hiệu quả giao hàng.\n")
md.append("![Bản đồ Quy hoạch Đà Lạt](file:///C:/Users/lap4all/.gemini/antigravity-ide/brain/2488f1f5-238f-4151-b3dd-1b59691a739c/images/da_lat_ban_do_quy_hoach.png)")
md.append("*Hình 1: Bản đồ Quy hoạch Mạng lưới Bưu cục Thành phố Đà Lạt năm 2026*\n")

md.append("#### 1.2. Thành phố Bảo Lộc")
md.append("- **Phường 1 Bảo Lộc (Mã: 24823):** Sản lượng AM trình bày: `1.217.0 đơn/ngày` (3.634.0 kg/ngày) | Giao: 618.0 · Lấy: 599.0. Hiện do (LDO) B'Lao (65.5%) và (LDO) 1 Bảo Lộc (34.5%) phụ trách. AM đề xuất GỘP VỀ BC CHÍNH `(LDO) B'Lao` (AM Hồng Bích Nga).")
md.append("- **Phường 2 Bảo Lộc (Mã: 24820):** Sản lượng AM trình bày: `789.0 đơn/ngày` (1.438.0 kg/ngày) | Giao: 454.0 · Lấy: 335.0. GỘP VỀ BC CHÍNH `(LDO) B'Lao` (AM Hồng Bích Nga).")
md.append("- **Phường B'Lao (Mã: 24829):** Sản lượng AM trình bày: `678.0 đơn/ngày` (1.341.0 kg/ngày) | Giao: 531.0 · Lấy: 147.0. GỘP VỀ BC CHÍNH `(LDO) 3 Bảo Lộc` (AM Nguyễn Lê Nguyên Vũ).")
md.append("- **Xã Bảo Lâm 2 (Mã: 25084):** Sản lượng AM trình bày: `324.0 đơn/ngày` (770.0 kg/ngày) | Giao: 274.0 · Lấy: 50.0. Điều chuyển về BC Bảo Lâm 3.\n")

md.append("#### 1.3. Khu vực Lâm Hà - Đam Rông")
md.append("- **Xã Nam Hà Lâm Hà (Mã: 24883):** Sản lượng AM trình bày: `104.0 đơn/ngày` | Giao: 101.0 · Lấy: 3.0. Chuyển phần xã cũ Phi Tô [60.0 đơn/ngày] từ BC Đinh Văn sang BC Nam Ban (AM Huỳnh Thị Kim Chi). BC Đinh Văn còn ~640 đơn/ngày (11 NV), BC Nam Ban nâng lên ~510 đơn/ngày (9 NV). AM tìm mặt bằng mới rộng hơn MB 96m2 hiện tại.")
md.append("- **Xã Đam Rông 4 (Mã: 24853):** Sản lượng AM trình bày: `59.0 đơn/ngày` | Giao: 56.0 · Lấy: 3.0. Currently split between BC Đam Rông 3 và BC Lang Biang - Đà Lạt 1.")
md.append("  - > [!IMPORTANT]")
md.append("    > **LÝ DO GIỮ NGUYÊN TUYẾN THÔN ĐƯNG KNỚ THUỘC BC LANG BIANG (AM Huỳnh Thị Kim Chi):** Thôn Đưng Knớ (sau sáp nhập thuộc Xã Đam Rông 4) hiện do BC Lang Biang - Đà Lạt 1 phụ trách với sản lượng ~10 đơn/ngày. Khoảng cách từ BC Đam Rông 3 đến Thôn Đưng Knớ là >50 km (phạm vi địa bàn rộng >30 km), đường núi rất xa và ít đơn. Do đó, AM đề xuất **GIỮ NGUYÊN** tuyến cover Thôn Đưng Knớ thuộc BC Lang Biang - Đà Lạt 1 quản lý như hiện tại.\n")
md.append("![Sơ đồ Tuyến Cover Lâm Hà](file:///C:/Users/lap4all/.gemini/antigravity-ide/brain/2488f1f5-238f-4151-b3dd-1b59691a739c/images/lam_ha_tuyen_cover.png)")
md.append("*Hình 2: Sơ đồ Tuyến cover hiện tại Khu vực Lâm Hà - Đam Rông*\n")
md.append("![Bản đồ Quy hoạch Lâm Hà Đam Rông](file:///C:/Users/lap4all/.gemini/antigravity-ide/brain/2488f1f5-238f-4151-b3dd-1b59691a739c/images/lam_ha_ban_do_quy_hoach.png)")
md.append("*Hình 3: Bản đồ Quy hoạch Khu vực Lâm Hà - Đam Rông theo ĐVHC Mới 2026*\n")

md.append("---")
md.append("### 2. Tỉnh Khánh Hòa\n")
md.append("#### 2.1. Thành phố Nha Trang")
md.append("- **Phường Nam Nha Trang (Mã: 22402):**")
md.append("  - **Sản lượng AM trình bày:** `2.418.0 đơn/ngày` (4.664.0 kg/ngày) | Giao: 1.340.0 · Lấy: 1.078.0. Hiện do 3 BC phụ trách: (KHO) Nam Nha Trang 3 (45.9% - 1.109 đơn), (KHO) Nam Nha Trang 1 (34.8% - 842 đơn), (KHO) Nam Nha Trang 5 (19.3% - 467 đơn). Sáp nhập từ Phường Phước Hải, Phước Long, Vĩnh Trường, Vĩnh Thái, Phước Đồng.")
md.append("  - > [!IMPORTANT]")
md.append("    > **LÝ DO GIỮ NGUYÊN 2 BƯU CỤC PHỤ TRÁCH (AM Thái Thị Thanh Thư):**")
md.append("    > 1. Sản lượng Phường Nam Nha Trang quá lớn (2.418 đơn/ngày), một bưu cục không thể tải nổi kho bãi và nhân sự.")
md.append("    > 2. Bưu cục (KHO) Nam Nha Trang 5 nằm tại Phước Đồng là khu vực xã xa, đường đi hiểm trở, địa bàn nằm biệt lập phía ngoài rìa thành phố giáp ranh đèo và Cam Lâm. Nếu gộp chung về cụm Nam Nha Trang trung tâm sẽ làm khoảng cách di chuyển quá xa (>20km đường đèo).")
md.append("    > => Phương án AM: Đóng cửa BC Nam Nha Trang 2 & Nam Nha Trang 3. Di dời BC Nam Nha Trang 1 ra vị trí trung tâm phường mới. **GIỮ NGUYÊN** BC Nam Nha Trang 5 phụ trách cụm xã biệt lập Phước Đồng.")
md.append("- **Phường Nha Trang (Mã: 22366):** Sản lượng AM trình bày: `1.592.0 đơn/ngày` (4.948.0 kg/ngày) | Giao: 1.205.0 · Lấy: 387.0. Di dời BC Nha Trang ra mặt bằng rộng ở trung tâm phường mới và gộp tuyến các BC cũ về BC Nha Trang duy nhất quản lý.")
md.append("- **Phường Tây Nha Trang (Mã: 22390):** Sản lượng AM trình bày: `747.0 đơn/ngày` | Giao: 610.0 · Lấy: 137.0. Gộp về BC (KHO) Tây Nha Trang (AM Phan Đình Duy).\n")
md.append("![Bản đồ Quy hoạch Nha Trang](file:///C:/Users/lap4all/.gemini/antigravity-ide/brain/2488f1f5-238f-4151-b3dd-1b59691a739c/images/nha_trang_ban_do_quy_hoach.png)")
md.append("*Hình 4: Bản đồ Quy hoạch Mạng lưới Bưu cục Thành phố Nha Trang 2026*\n")

md.append("#### 2.2. Huyện Diên Khánh & Huyện Vạn Ninh")
md.append("- **Xã Diên Khánh (Mã: 22651):** Sản lượng AM trình bày: `425.0 đơn/ngày`. Gộp về BC Diên Khánh 2 (AM Nguyễn Hoàng Phi). BC Diên Khánh 1 phụ trách Vĩnh Thạnh (170 đơn), Vĩnh Trung (110 đơn), Diên An (120 đơn).")
md.append("- **Xã Vạn Thắng (Mã: 22516):** Sản lượng AM trình bày: `211.0 đơn/ngày` (475.0 kg/ngày) | Giao: 205.0 · Lấy: 6.0. Hiện do BC Tu Bông (72.5%) và BC Vạn Ninh (27.5%) phụ trách.")
md.append("  - > [!NOTE]")
md.append("    > **LÝ DO GỘP VỀ BC TU BÔNG (AM Phạm Bá Thành Công):** Bưu cục Tu Bông nằm đúng trục địa giới hành chính mới, vị trí kho bãi phù hợp với phạm vi giao nhận của 2 xã cũ (Vạn Thắng + Vạn Bình) và có diện tích kho đủ lớn cho vận hành dài hạn.\n")
md.append("![Bản đồ Quy hoạch Vạn Ninh](file:///C:/Users/lap4all/.gemini/antigravity-ide/brain/2488f1f5-238f-4151-b3dd-1b59691a739c/images/van_ninh_ban_do_quy_hoach.jpg)")
md.append("*Hình 5: Bản đồ Quy hoạch Bưu cục Khu vực Vạn Ninh 2026*\n")

md.append("---")
md.append("### 3. Tỉnh Ninh Thuận\n")
md.append("- **Phường Phan Rang (Mã: 22759):** Sản lượng AM trình bày: `802.0 đơn/ngày` | Giao: 509.0 · Lấy: 293.0. Gộp về BC chính (NTH) Phan Rang.")
md.append("- **Phường Ninh Chử (Mã: 22834):** Sản lượng AM trình bày: `466.0 đơn/ngày` (1.079.0 kg/ngày) | Giao: 299.0 · Lấy: 167.0. Gộp về BC chính (NTH) Phan Rang.")
md.append("- **Xã Ninh Hải (Mã: 22852):** Sản lượng AM trình bày: `185.0 đơn/ngày` (376.0 kg/ngày) | Giao: 178.0 · Lấy: 7.0. Gộp về BC (NTH) Ninh Chử.")
md.append("- **Mở mới 01 Bưu cục:** `BC (NTH) Đông Hải` phụ trách 6 phường ven biển (Đông Hải, Mỹ Bình, Mỹ Đông, Mỹ Hải, Đạo Long, Kinh Dinh). Sản lượng: Giao 600 đơn/ngày, Lấy 250 đơn/ngày. Nhân sự: 1 NVXL - 7 NVPTTT.\n")
md.append("![Bản đồ Quy hoạch Ninh Thuận](file:///C:/Users/lap4all/.gemini/antigravity-ide/brain/2488f1f5-238f-4151-b3dd-1b59691a739c/images/ninh_thuan_ban_do_quy_hoach.png)")
md.append("*Hình 6: Bản đồ Quy hoạch Mạng lưới Bưu cục Tỉnh Ninh Thuận 2026*\n")

md.append("---")
md.append("### 4. Tỉnh Bình Thuận\n")
md.append("#### 4.1. Cụm Hàm Thuận Bắc & Hàm Liêm")
md.append("- **Phường Hàm Thắng (Mã: 22933):** Sản lượng AM trình bày: `618.0 đơn/ngày` (1.939.0 kg/ngày) | Giao: 545.0 · Lấy: 73.0. Gộp về BC (BTH) Phú Thủy (AM Nguyễn Ngọc Khánh).")
md.append("- **Phường Phan Thiết (Mã: 22945):** Sản lượng AM trình bày: `905.0 đơn/ngày` (2.112.0 kg/ngày) | Giao: 777.0 · Lấy: 128.0. Gộp về BC (BTH) Hàm Thắng (AM Nguyễn Ngọc Khánh).")
md.append("- **Phường Bình Thuận (Mã: 22960):** Sản lượng AM trình bày: `680.0 đơn/ngày` (1.577.0 kg/ngày) | Giao: 611.0 · Lấy: 69.0. Gộp về BC (BTH) Hàm Thắng (AM Nguyễn Ngọc Khánh).\n")
md.append("![Bản đồ Quy hoạch Hàm Thuận Bắc](file:///C:/Users/lap4all/.gemini/antigravity-ide/brain/2488f1f5-238f-4151-b3dd-1b59691a739c/images/ham_thuan_bac_ban_do.jpg)")
md.append("*Hình 7: Bản đồ Quy hoạch Cụm Bưu cục Hàm Thuận Bắc - Hàm Liêm 2026*\n")
md.append("![Bản đồ Quy hoạch Phú Thủy Phan Thiết](file:///C:/Users/lap4all/.gemini/antigravity-ide/brain/2488f1f5-238f-4151-b3dd-1b59691a739c/images/phu_thuy_phan_thiet_ban_do.jpg)")
md.append("*Hình 8: Bản đồ Quy hoạch Cụm Bưu cục Phú Thủy - Hàm Thắng - Phan Thiết 2026*\n")

md.append("#### 4.2. Cụm Tánh Linh - Đức Linh")
md.append("- **Mở mới 01 Bưu cục:** `BC (BTH) Nam Thành` (Cover Nam Thành 250 đơn, Nghị Đức 200 đơn; Giao 450-500, Lấy 50-60; Nhân sự: 7 NV, 1 NVXL).\n")
md.append("![Bản đồ Quy hoạch Tánh Linh Đức Linh](file:///C:/Users/lap4all/.gemini/antigravity-ide/brain/2488f1f5-238f-4151-b3dd-1b59691a739c/images/tanh_linh_duc_linh_ban_do.jpg)")
md.append("*Hình 9: Bản đồ Quy hoạch Bưu cục Tánh Linh - Đức Linh 2026*\n")

md.append("---")
md.append("### 5. Tỉnh Đắc Nông\n")
md.append("> [!IMPORTANT]")
md.append("> **LÝ DO GIỮ NGUYÊN 100% BƯU CỤC KHÔNG GỘP TẠI ĐẮK NÔNG (AM Trần Thị Nhung & AM Trần Văn Phước):**")
md.append("> - **Các xã rà soát:** **Phường Bắc Gia Nghĩa** (Mã: 24611 - 734 đơn), **Phường Nam Gia Nghĩa** (Mã: 24615 - 499 đơn), **Xã Tà Đùng** (Mã: 24637 - 187 đơn), **Xã Đắc Sắk** (Mã: 24678 - 194 đơn), **Xã Đức An** (Mã: 24717 - 327 đơn), **Xã Quảng Tân** (Mã: 24748 - 201 đơn).")
md.append("> 1. **Rủi ro nhân sự 100%:** 100% nhân sự hiện tại tại các điểm Bưu cục xã ven khẳng định KHÔNG THEO BC MỚI nếu sáp nhập cưỡng ép do địa hình miền núi chia cắt, khoảng cách di chuyển quá xa. Việc gộp đường đột sẽ dẫn đến nghỉ việc hàng loạt và sập toàn bộ tuyến giao.")
md.append("> 2. **Địa lý chia cắt:** Như Xã Đắc Sắk (76.3% do BC Đức Lập đảm nhận, phần Nam Xuân do BC Krông Nô), Xã Đức An (phần Đắk N'Drung gần BC Trường Xuân hơn BC Đức An).")
md.append("> 3. **Tham khảo đối thủ:** 100% đối thủ trong ngành (GHTK, Viettel Post, Shopee Express, J&T) đều TẠM THỜI GIỮ NGUYÊN Bưu cục theo xã cũ, chưa đơn vị nào gộp tuyến theo xã mới.")
md.append("> => **Đề xuất AM:** Kính đề nghị Ban Lãnh đạo cho phép **TẠM THỜI GIỮ NGUYÊN 100% Bưu cục** và phạm vi quản lý theo địa giới cũ tại Đắk Nông trong 6 tháng tới.\n")

md.append("---")
md.append("## IV. BẢNG TỔNG HỢP BIẾN ĐỘNG MẠNG LƯỚI BƯU CỤC TOÀN VÙNG NTB\n")
md.append("| Loại Biến Động | Tên Bưu Cục | Tỉnh / Khu Vực | Chi Tiết Phương Án & Bố Trí Nhân Sự |")
md.append("| :--- | :--- | :--- | :--- |")
md.append('| <font color="green">**MỞ MỚI**</font> | BC (LDO) Xuân Hương - Đà Lạt 2 | Lâm Đồng (Đà Lạt) | Mở mới tại Phường 10 phụ trách Phường 3 + Phường 10 cũ. |')
md.append('| <font color="green">**MỞ MỚI**</font> | BC (NTH) Đông Hải | Ninh Thuận (Phan Rang) | Mở mới phụ trách 6 phường coastal, 1 NVXL - 7 NVPTTT. |')
md.append('| <font color="green">**MỞ MỚI**</font> | BC (BTH) Nam Thành | Bình Thuận (Tánh Linh) | Mở mới phụ trách Nam Thành & Nghị Đức, 1 NVXL - 7 NVPTTT. |')
md.append('| <font color="red">**ĐÓNG CỬA**</font> | BC (LDO) 1 Bảo Lộc | Lâm Đồng (Bảo Lộc) | Đóng cửa, gộp toàn bộ tuyến và nhân sự về BC B\'Lao. |')
md.append('| <font color="red">**ĐÓNG CỬA**</font> | BC (KHO) Nam Nha Trang 2 | Khánh Hòa (Nha Trang) | Đóng cửa, gộp tuyến về BC Nam Nha Trang 1 & Nam Nha Trang 5. |')
md.append('| <font color="red">**ĐÓNG CỬA**</font> | BC (KHO) Nam Nha Trang 3 | Khánh Hòa (Nha Trang) | Đóng cửa, gộp tuyến về BC Nam Nha Trang 1 & Nam Nha Trang 5. |')
md.append('| <font color="blue">**DI DỜI / GIỮ**</font> | BC (KHO) Nha Trang | Khánh Hòa (Nha Trang) | Di dời ra mặt bằng rộng hơn ở trung tâm Phường Nha Trang mới. |')
md.append('| <font color="blue">**DI DỜI / GIỮ**</font> | BC (KHO) Nam Nha Trang 1 | Khánh Hòa (Nha Trang) | Di dời ra vị trí trung tâm Phường Nam Nha Trang mới. |')
md.append('| <font color="blue">**DI DỜI / GIỮ**</font> | BC (LDO) Nam Ban Lâm Hà | Lâm Đồng (Lâm Hà) | Tìm mặt bằng rộng hơn MB 96m2 hiện tại để gộp tuyến xã Phi Tô. |')
md.append('| <font color="blue">**DI DỜI / GIỮ**</font> | BC (BTH) Hàm Thuận & Hàm Liêm | Bình Thuận (Hàm Thuận B.) | Di dời vị trí kho bãi phù hợp với địa giới hành chính mới. |\n')

md.append("---")
md.append("## V. LỘ TRÌNH TRIỂN KHAI & ĐỀ XUẤT BAN LÃNH ĐẠO\n")
md.append("1. **Giai đoạn 1 (Tháng 1 - 2/2026):** Điều chuyển lập tức 18 tuyến giao chéo xa ranh giới (30-40km) về các Bưu cục lân cận gần hơn để giảm chi phí di chuyển.")
md.append("2. **Giai đoạn 2 (Tháng 3 - 4/2026):** Tiến hành khảo sát và ký hợp đồng mặt bằng di dời cho các BC Nha Trang, Nam Nha Trang 1, Nam Ban, Hàm Thuận, Hàm Liêm; Đàm phán thanh lý HĐ mặt bằng các BC đóng cửa (Nam Nha Trang 2, 3, 1 Bảo Lộc).")
md.append("3. **Giai đoạn 3 (Tháng 5 - 6/2026):** Khai trương 03 Bưu cục mở mới (Xuân Hương 2, Đông Hải, Nam Thành); Hoàn tất việc điều chuyển nhân sự và chính thức áp dụng ranh giới ĐVHC mới trên hệ thống.\n")

md.append("> [!TIP]")
md.append("> **Kết luận:** Kính đề nghị Ban Lãnh đạo xem xét phê duyệt Kế hoạch Quy hoạch Mạng lưới Bưu cục Vùng NTB 2026 để Khối Vận hành và các AM chủ động triển khai theo lộ trình.")

with open(md_path, "w", encoding="utf-8") as f:
    f.write("\n".join(md))

print("Markdown artifact updated successfully with AM volumes and Giu Nguyen rationale!")
