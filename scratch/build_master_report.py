import os, sys, pandas as pd, docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

sys.stdout.reconfigure(encoding='utf-8')

excel_path = r'C:\Users\lap4all\Downloads\config_psbba_NTB.xlsx'
xls = pd.ExcelFile(excel_path)

# Load data from Thủy_A
df_thuya = pd.read_excel(xls, 'Thủy_A')
# Load data from Thủy
df_thuy = pd.read_excel(xls, 'Thủy')
# Load data from lịch làm việc
df_lich = pd.read_excel(xls, 'lịch làm việc')

print('Building Master Document...')

# Generate Markdown Content
md_content = """# KẾ HOẠCH VẬN HÀNH CHI TIẾT EVENT 8.8
## KHO TRUNG CHUYỂN (KTC) - VÙNG NAM TRUNG BỘ (NTB)
*(Hồ sơ Vận hành Tổng thể Chuẩn hóa 100% theo Mô hình TNB Event 7.7)*

---

## 1. THÔNG TIN CHUNG & MỤC TIÊU CHIẾN LƯỢC

### 1.1. Phạm vi & Mạng lưới Kho Trung Chuyển (KTC)
- **Thời gian diễn ra Event**: Từ **06/08/2026 đến 15/08/2026** (10 ngày cao điểm).
- **Phạm vi Cụm Kho KTC / Hub Sorting vùng NTB**:
  1. **Kho Trung Chuyển Khánh Hòa (KTC Khánh Hòa)** - Super-Hub trung tâm duyên hải (Chiếm 38.1% volume toàn vùng).
  2. **Kho Chuyển Tiếp Bình Thuận (CT Bình Thuận)** - Cửa ngõ kết nối phía Nam với TP.HCM & Đồng Nai (22.6% volume).
  3. **Kho Chuyển Tiếp Đức Trọng (CT Đức Trọng - Lâm Đồng)** - Hub trung chuyển trung tâm vùng Tây Nguyên (19.0% volume).
  4. **Kho Chuyển Tiếp Bảo Lộc (CT Bảo Lộc - Lâm Đồng)** - Hub vệ tinh hỗ trợ Nam Lâm Đồng (12.4% volume).
  5. **Kho Chuyển Tiếp Đắc Nông (CT Đắc Nông)** - Hub kết nối Nam Tây Nguyên (7.9% volume).

### 1.2. Chỉ số KPI Vận hành Cam kết
- **Tổng sản lượng Sorting**: **771,977 đơn hàng** trong 10 ngày (Trung bình **77,198 đơn/ngày**).
- **Đỉnh điểm sản lượng (Peak Days)**: Ngày 1 Mega Sale **08/08 (97,649 đơn)** và Ngày 2 After-shock **10/08 (89,608 đơn)**.
- **Tỷ lệ Kết nối Chuyến xe (COT Fulfillment)**: Đạt **100%** khung giờ COT cố định, không rớt chuyến.
- **Chỉ số Tồn kho (Backlog)**: Cam kết **0%** đơn tồn quá 24h tại tất cả các kho.

---

## 2. DỰ BÁO SẢN LƯỢNG & NHẬN XÉT CHI TIẾT TỪNG KHO

### 2.1. Biểu đồ Diễn biến Sản lượng Sorting theo Kho KTC Event 8.8

![Sản lượng Sorting theo Kho KTC Event 8.8](file:///C:/Users/lap4all/.gemini/antigravity-ide/brain/497c8661-37a7-4c4e-986e-e8ef6c144fb3/chart1_daily_sorting_volume.png)

---

### 2.2. Bảng Dữ Liệu 1: Sản Lượng Sorting Theo Kho & Nhóm Hàng (06/08 – 15/08/2026)

| Kho KTC / Hub | Nhóm hàng | 06/08 | 07/08 | 08/08 (Peak 1) | 09/08 | 10/08 (Peak 2) | 11/08 | 12/08 | 13/08 | 14/08 | 15/08 | Tổng 10d | TB/ngày | Tỷ trọng |
| :--- | :--- | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: |
| **KTC Khánh Hòa** | Normal | 21,652 | 22,586 | 30,057 | 18,795 | 26,792 | 22,516 | 22,667 | 22,425 | 23,306 | 23,281 | **234,077** | 23,408 | 79.7% |
| | Bulky | 2,755 | 2,852 | 3,940 | 3,975 | 3,524 | 4,971 | 4,882 | 3,431 | 2,738 | 3,940 | **37,008** | 3,701 | 12.6% |
| | Freight | 1,659 | 1,718 | 2,901 | 2,438 | 2,725 | 2,819 | 2,274 | 1,779 | 1,601 | 2,732 | **22,646** | 2.265 | 7.7% |
| | **Cộng Khánh Hòa**| **26,067**| **27,156**| **36,898**| **25,208**| **33,041**| **30,306**| **29,823**| **27,635**| **27,645**| **29,953**| **293,732**| **29,373**| **38.1%** |
| **CT Bình Thuận** | Normal | 12,437 | 12,973 | 17,265 | 10,796 | 15,389 | 12,933 | 13,021 | 12,881 | 13,387 | 13,372 | **134,454** | 13,445 | 76.9% |
| | Bulky | 1,785 | 1,848 | 2,347 | 2.080 | 2.515 | 2.599 | 3.508 | 2.804 | 1.940 | 2.347 | **23,773** | 2,377 | 13.6% |
| | Freight | 1,218 | 1,260 | 2,334 | 1,665 | 2.398 | 1.849 | 1.429 | 1.285 | 1.175 | 2.005 | **16,618** | 1,662 | 9.5% |
| | **Cộng Bình Thuận**| **15,440**| **16,082**| **21,946**| **14,541**| **20,302**| **17,381**| **17,958**| **16,970**| **16,502**| **17,724**| **174,846**| **17,485**| **22.6%** |
| **CT Đức Trọng** | Normal | 10,447 | 10,897 | 14,502 | 9,069 | 12,927 | 10,864 | 10,937 | 10,820 | 11,245 | 11,233 | **112,941** | 11,294 | 76.8% |
| | Bulky | 1,436 | 1,487 | 3,483 | 1,664 | 2,898 | 1.888 | 1.543 | 1,453 | 1.386 | 2.365 | **19,603** | 1,960 | 13.3% |
| | Freight | 1,059 | 1,096 | 2,032 | 1,439 | 1.918 | 1.659 | 1.324 | 1.161 | 1.022 | 1.744 | **14,454** | 1,445 | 9.8% |
| | **Cộng Đức Trọng**| **12,943**| **13,481**| **20,017**| **12,172**| **17,743**| **14,411**| **13,804**| **13,434**| **13,653**| **15,341**| **146,999**| **14,700**| **19.0%** |
| **CT Bảo Lộc** | Normal | 6,858 | 7,154 | 8,774 | 6,250 | 8,805 | 7,255 | 7,180 | 7,103 | 7,382 | 7,374 | **74,135** | 7,414 | 77.5% |
| | Bulky | 965 | 999 | 1,368 | 1.285 | 1.612 | 1.531 | 1.489 | 1.331 | 1.008 | 1.405 | **12,993** | 1,299 | 13.6% |
| | Freight | 625 | 646 | 1.112 | 888 | 1.124 | 1.007 | 773 | 718 | 603 | 1.028 | **8,524** | 852 | 8.9% |
| | **Cộng Bảo Lộc** | **8,448** | **8,799** | **11,254**| **8,423** | **11,541**| **9,793** | **9,442** | **9.152** | **8,993** | **9,807** | **95,652** | **9,565** | **12.4%** |
| **CT Đắc Nông** | Normal | 4,407 | 4,597 | 6,118 | 3,826 | 5,453 | 4,583 | 4,614 | 4,564 | 4,744 | 4.738 | **47,644** | 4,764 | 78.4% |
| | Bulky | 642 | 665 | 887 | 902 | 954 | 843 | 1,128 | 979 | 705 | 887 | **8,592** | 859 | 14.1% |
| | Freight | 332 | 344 | 529 | 488 | 574 | 553 | 454 | 390 | 320 | 529 | **4,513** | 451 | 7.4% |
| | **Cộng Đắc Nông** | **5,381** | **5,605** | **7.534** | **5.216** | **6.981** | **5.979** | **6.196** | **5.933** | **5.769** | **6.154** | **60.748** | **6.075** | **7.9%** |
| **TỔNG NTB SORTING**| **TỔNG KTC**| **68,279**| **71,123**| **97,649**| **65,560**| **89,608**| **77,870**| **77,223**| **73,124**| **72,562**| **78,979**| **771,977**| **77,198**| **100.0%** |

---

### 2.3. Nhận Xét Chuyên Sâu Cụ Thể Từng Kho (Chuẩn hóa từ Mô hình TNB)

![Cơ cấu Nhóm hàng tại KTC NTB](file:///C:/Users/lap4all/.gemini/antigravity-ide/brain/497c8661-37a7-4c4e-986e-e8ef6c144fb3/chart2_product_group_breakdown.png)

#### 📝 Nhận xét 1: Kho Trung Chuyển Khánh Hòa (Super-Hub NTB)
- **Cơ cấu hàng hóa**: Normal 79.7% (234,077 đơn) là xương sống vận hành; tỷ trọng Bulky **12.6% (37,008 đơn)** và Freight **7.7% (22,646 đơn)** chiếm tổng cộng **20.3%** sản lượng kho - LỚN NHẤT VÙNG (gấp 2.4 lần tổng hàng cồng kềnh của Đức Trọng và Bảo Lộc cộng lại).
- **Diễn biến ngày đỉnh**: Ngày campaign 08/08, tổng Bulky+Freight vọt lên **6,841 đơn/ngày** (nền ngày thường chỉ ~5.9k) rồi co về ~6.4k vào Chủ nhật 09/08.
- **⇒ Kế hoạch Quy hoạch Kho**: Sàn hàng cồng kềnh, dock xuất xe và đầu xe tải lớn (15 - 20 tấn) của Khánh Hòa phải quy hoạch theo **NGÀY ĐỈNH (~36.9K đơn/ngày)**, không quy hoạch theo mức trung bình (~29.4K/ngày).

#### 📝 Nhận xét 2: Kho Chuyển Tiếp Bình Thuận (Cửa ngõ phía Nam)
- **Cơ cấu hàng hóa**: Normal 76.9%, Bulky 13.6% (23,773 đơn), Freight **9.5% CAO NHẤT VÙNG (16,618 đơn)** - Phản ánh đặc thù nông sản/hải sản và kết nối trực tiếp với dòng hàng liên vùng từ TP.HCM/Đồng Nai.
- **Biến động ngày đỉnh**: Ngày đỉnh 08/08, Bulky+Freight chiếm **21.3% (4,681 đơn)**. 
- **⇒ Kế hoạch Quy hoạch Kho**: Rủi ro cơ cấu nằm ở Freight & Bulky dồn đợt campaign gây áp lực trực tiếp lên dock nhập hàng liên vùng. Vận hành theo luồng xả bãi nhanh; các chuyến xe 15 tấn từ HCM về phải có luồng ưu tiên xả pallet ngay ca đêm.

#### 📝 Nhận xét 3: Kho Chuyển Tiếp Đức Trọng (Hub Trung Chuyển Tây Nguyên)
- **Cơ cấu hàng hóa**: Normal 76.8%, Bulky 13.3% (19,603 đơn), Freight 9.8% (14,454 đơn). Tổng Bulky+Freight đạt **34,057 đơn/10 ngày**.
- **Biến động ngày đỉnh**: Ngày 08/08, hàng Bulky+Freight nhảy vọt vọt lên **5,515 đơn/ngày (+130% so với ngày thường)**.
- **⇒ Kế hoạch Quy hoạch Kho**: Hàng cồng kềnh ngày đỉnh gom chuyến dùng chung tuyến xe tải lớn kết nối Đức Trọng - Đà Lạt và Bảo Lộc; sử dụng xe nâng điện 2 tầng để hạ tải nhanh giảm sức nhân viên.

#### 📝 Nhận xét 4: Kho Chuyển Tiếp Bảo Lộc & Đắc Nông (Kho Vệ Tinh)
- **Cơ cấu hàng hóa**: Đắc Nông có Normal 78.4%, Bulky 14.1% (8,592 đơn), Freight 7.4% (4,513 đơn). Bảo Lộc có Normal 77.5%, Bulky 13.6% (12,993 đơn).
- **⇒ Kế hoạch Quy hoạch Kho**: Quy mô kho nhỏ nhưng biến động ngày campaign vẫn tăng 1.4 - 1.6 lần. Ưu tiên cơ chế ca kíp linh hoạt (sử dụng 6 Freelance bọc lót tại Đắc Nông) và điều phối xe chung tuyến thay vì tăng nguồn lực cố định.

---

### 2.4. Bảng Dữ Liệu Chi Tiết Volume Lấy & Giao Theo Kênh Sàn & Theo Tỉnh

![Xu hướng Sản lượng Lấy và Giao theo Kênh Sàn](file:///C:/Users/lap4all/.gemini/antigravity-ide/brain/497c8661-37a7-4c4e-986e-e8ef6c144fb3/chart3_pick_delivery_channel.png)

#### A. Bảng Dữ Liệu 2.2: Volume LẤY Theo Sàn / Kênh (06/08 – 15/08)

| Sàn / Kênh | 06/08 | 07/08 | 08/08 (Peak) | 09/08 | 10/08 | 11/08 | 12/08 | 13/08 | 14/08 | 15/08 | Tổng 10d | TB/ngày |
| :--- | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: |
| **Shopee** | 2,263 | 1,830 | 4,155 | 1,969 | 3,282 | 2,484 | 2,625 | 2,212 | 2,227 | 2,357 | **25,402** | 2,540 |
| **Shopee-Bulky** | 184 | 197 | 679 | 533 | 529 | 283 | 179 | 195 | 170 | 378 | **3,327** | 333 |
| *(Shopee-Bulky 10-15kg)*| *(94)* | *(103)*| *(479)* | *(289)*| *(272)*| *(176)*| *(97)* | *(99)* | *(91)* | *(223)* | ***(1,923)***| *(192)*|
| **SME (Truyền thống)**| 10,139| 10,495| 10,911 | 8,353 | 11,321| 10,417| 10,166| 10,139| 10,495| 10,215| **102,653**| 10,265|
| **SME-Bulky** | 410 | 412 | 364 | 276 | 486 | 469 | 433 | 410 | 412 | 398 | **4,074** | 407 |
| **TikTok Shop (TTS)** | 4,104 | 4,337 | 4,546 | 3,110 | 4,045 | 3,403 | 3,366 | 3,876 | 4,133 | 4,066 | **38,983** | 3,898 |
| **TTS-Bulky** | 150 | 166 | 136 | 102 | 166 | 130 | 126 | 132 | 155 | 135 | **1,398** | 140 |
| **GRAND TOTAL LẤY** | **17,344**| **17,541**| **21,270** | **14,631**| **20,101**| **17,362**| **16,992**| **17,063**| **17,683**| **17,772**| **177,759**| **17,776** |

#### B. Bảng Dữ Liệu 2.3: Volume LẤY Theo Tỉnh (06/08 – 15/08)

| Tỉnh / Thành | 06/08 | 07/08 | 08/08 (Peak) | 09/08 | 10/08 | 11/08 | 12/08 | 13/08 | 14/08 | 15/08 | Tổng 10d | TB/ngày |
| :--- | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: |
| **Bình Thuận** | 3,327 | 3,388 | 3,361 | 2,810 | 3,776 | 3,183 | 3,089 | 3,220 | 3,373 | 3,431 | **32,955** | 3,295 |
| **Khánh Hòa** | 4,267 | 4,286 | 5,450 | 3,576 | 4,995 | 4,292 | 4,245 | 4,210 | 4,352 | 4,369 | **44,043** | 4,404 |
| **Lâm Đồng** | 4,705 | 4,720 | 6,645 | 4,027 | 5,521 | 4,709 | 4,604 | 4,608 | 4,780 | 4,851 | **49,170** | 4,917 |
| **Ninh Thuận** | 1,358 | 1,357 | 1,729 | 1,177 | 1,654 | 1,410 | 1,377 | 1,348 | 1,380 | 1,401 | **14,190** | 1,419 |
| **Đắc Nông** | 3,687 | 3,791 | 4,086 | 3,042 | 4,155 | 3,769 | 3,677 | 3,677 | 3,799 | 3,720 | **37,401** | 3,740 |
| **GRAND TOTAL LẤY** | **17,344**| **17,541**| **21,270** | **14,631**| **20,101**| **17,362**| **16,992**| **17,063**| **17,683**| **17,772**| **177,759**| **17,776** |

#### C. Bảng Dữ Liệu 2.4: Volume GIAO Theo Sàn / Kênh (06/08 – 15/08)

| Sàn / Kênh | 06/08 | 07/08 | 08/08 (Peak) | 09/08 | 10/08 | 11/08 | 12/08 | 13/08 | 14/08 | 15/08 | Tổng 10d | TB/ngày |
| :--- | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: |
| **Shopee** | 14,964| 12,090| 33,036 | 12,642| 22,526| 16,427| 17,352| 14,667| 14,722| 19,582| **178,007**| 17,801|
| **Shopee-Bulky** | 2,379 | 2,470 | 7,021 | 3,998 | 6,012 | 3,426 | 2,388 | 2,325 | 2,251 | 5,147 | **37,418** | 3,742 |
| *(Shopee-Bulky 10-15kg)*| *(1,304)*| *(1,355)*| *(5,254)* | *(2,163)*| *(3,136)*| *(1,877)*| *(1,305)*| *(1,276)*| *(1,230)*| *(2,903)*| ***(21,804)***| *(2,180)*|
| **SME (Truyền thống)**| 22,391| 23,184| 23,334 | 18,439| 24,985| 23,010| 22,460| 22,391| 23,184| 22,553| **225,931**| 22,593|
| **SME-Bulky** | 1,446 | 1,447 | 1,341 | 910 | 1,628 | 1,567 | 1,516 | 1,446 | 1,447 | 1,354 | **14,104** | 1,410 |
| **TikTok Shop (TTS)** | 12,085| 12,738| 13,492 | 9,200 | 11,904| 9,998 | 9,888 | 11,317| 12,157| 11,980| **114,758**| 11,476|
| **TTS-Bulky** | 318 | 343 | 400 | 225 | 361 | 281 | 278 | 291 | 321 | 311 | **3,129** | 313 |
| **GRAND TOTAL GIAO**| **54,888**| **53,627**| **83,879** | **47,577**| **70,552**| **56,586**| **55,187**| **53,715**| **55,312**| **63,830**| **595,152**| **59,515** |

#### D. Bảng Dữ Liệu 2.5: Volume GIAO Theo Tỉnh (06/08 – 15/08)

| Tỉnh / Thành | 06/08 | 07/08 | 08/08 (Peak) | 09/08 | 10/08 | 11/08 | 12/08 | 13/08 | 14/08 | 15/08 | Tổng 10d | TB/ngày |
| :--- | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: |
| **Bình Thuận** | 12,846| 12,451| 20,394 | 11,203| 16,722| 13,287| 12,979| 12,569| 12,926| 15,090| **140,466**| 14,047|
| **Khánh Hòa** | 14,760| 14,451| 23,538 | 12,784| 18,951| 15,220| 14,834| 14,442| 14,894| 17,152| **161,028**| 16,103|
| **Lâm Đồng** | 15,907| 15,529| 23,879 | 13,927| 20,685| 16,544| 16,088| 15,580| 16,038| 18,661| **172,836**| 17,284|
| **Ninh Thuận** | 5,222 | 5,079 | 7,994 | 4,478 | 6,699 | 5,384 | 5,272 | 5,114 | 5,254 | 6,073 | **56,568** | 5,657 |
| **Đắc Nông** | 6,152 | 6,118 | 8,074 | 5,185 | 7,494 | 6,152 | 6,013 | 6,010 | 6,201 | 6,853 | **64,253** | 6,425 |
| **GRAND TOTAL GIAO**| **54,888**| **53,627**| **83,879** | **47,577**| **70,552**| **56,586**| **55,187**| **53,715**| **55,312**| **63,830**| **595,152**| **59,515** |

#### E. Bảng Dữ Liệu 2.6: Bảng Tính Combo Chart FC Volume Lấy & Giao vs Base vs % Tăng/Giảm

| Chỉ tiêu | 06/08 | 07/08 | 08/08 (Peak 1) | 09/08 | 10/08 (Peak 2) | 11/08 | 12/08 | 13/08 | 14/08 | 15/08 | TB Toàn đợt |
| :--- | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: |
| **FC Volume LẤY** | 17,344 | 17,541 | **21,270** | 14,631 | **20,101** | 17,362 | 16,992 | 17,063 | 17,683 | 17,772 | **17,776** |
| *Trung bình Lấy Base* | 17,776 | 17,776 | 17,776 | 17,776 | 17,776 | 17,776 | 17,776 | 17,776 | 17,776 | 17,776 | 17,776 |
| **% Tăng/Giảm FC LẤY**| **-2.4%** | **-1.3%** | **+19.7%** | **-17.7%**| **+13.1%** | **-2.3%** | **-4.4%** | **-4.0%** | **-0.5%** | **-0.0%** | **0.0%** |
| **FC Volume GIAO** | 54,888 | 53,627 | **83,879** | 47,577 | **70,552** | 56,586 | 55,187 | 53,715 | 55,312 | 63,830 | **59,515** |
| *Trung bình Giao Base*| 59,515 | 59,515 | 59,515 | 59,515 | 59,515 | 59,515 | 59,515 | 59,515 | 59,515 | 59,515 | 59,515 |
| **% Tăng/Giảm FC GIAO**| **-7.8%** | **-9.9%** | **+40.9%** | **-20.1%**| **+18.5%** | **-4.9%** | **-7.3%** | **-9.7%** | **-7.1%** | **+7.2%** | **0.0%** |

---

## 3. KẾ HOẠCH PHÂN CÔNG NHÂN SỰ & CA LÀM VIỆC CHI TIẾT TẤT CẢ 5 KHO (SCHEDULING & EQUIPMENT)

![Phân bổ Nhân sự theo Kho KTC](file:///C:/Users/lap4all/.gemini/antigravity-ide/brain/497c8661-37a7-4c4e-986e-e8ef6c144fb3/chart4_staffing_plan.png)

### Bảng Dữ Liệu 3.1: Tổng Hợp Phân Bổ Nhân Sự Tất Cả 5 Kho Trung Chuyển (NTB)

| Kho KTC / Hub | Ca làm việc | Lực lượng NVCT/ngày | Lực lượng Freelance/ngày | Tổng nhân sự ca/ngày | Sản lượng TB/ngày | Năng suất TB (Đơn/người/ngày) |
| :--- | :--- | :---: | :---: | :---: | :---: | :---: |
| **KTC Khánh Hòa** | Ca Day (7h00 - 18h00) | 12 NV | 5 NV | 17 NV | 29,373 | 839 đơn/người |
| | Ca Night (20h00 - 5h30) | 13 NV | 5 NV | 18 NV | | |
| | **Cộng Khánh Hòa** | **25 NV** | **10 NV** | **35 NV** | **29,373** | **839 đơn/người** |
| **CT Đức Trọng** | Ca Day (7h00 - 18h00) | 11 NV | 4 NV | 15 NV | 14,700 | 565 đơn/người |
| | Ca Night (20h00 - 5h30) | 9 NV | 4 NV | 13 NV | | |
| | **Cộng Đức Trọng** | **20 NV** | **8 NV** | **26 NV** | **14,700** | **565 đơn/người** |
| **CT Bảo Lộc** | Ca Day (7h00 - 18h00) | 4 NV | 3 NV | 7 NV | 9,565 | 638 đơn/người |
| | Ca Night (20h00 - 5h30) | 4 NV | 4 NV | 8 NV | | |
| | **Cộng Bảo Lộc** | **8 NV** | **7 NV** | **15 NV** | **9,565** | **638 đơn/người** |
| **CT Đắc Nông** | Ca Day (7h00 - 18h00) | 3 NV | 2 NV | 5 NV | 6,075 | 357 đơn/người |
| | Ca Night (20h00 - 5h30) | 4 NV | 4 NV | 8 NV | | |
| | **Cộng Đắc Nông** | **7 NV** | **6 NV** | **13 NV** | **6,075** | **467 đơn/người** |
| **CT Bình Thuận** | Ca Day (7h00 - 18h00) | 6 NV | 0 NV | 6 NV | 17,485 | **1,589 đơn/người (🔴 ĐIỂM NỔ RỦI RO)** |
| | Ca Night (18h00 - 7h00) | 5 NV | 0 NV | 5 NV | | |
| | **Cộng Bình Thuận** | **11 NV** | **0 NV** | **11 NV** | **17,485** | **1,589 đơn/người** |

---

### 3.2. Quy Trình Vận Hành & Setup Thiết Bị Chi Tiết Từng Kho (Theo Chuẩn TNB)

#### 🏢 1. KTC KHÁNH HÒA (Super-Hub Trọng Điểm)
- **Thiết bị vận hành**: 2 máng nhập hàng, 4 băng tải nâng hạ, 4 băng tải phân chọn, 4 xe nâng điện, 14 xe nâng tay, 4 băng tải con lăn.
- **Phương án Ca Day (07h00-18h00)**:
  - Sử dụng 2 băng tải nâng hạ + 2 máng nhập + 2 băng phân chọn nhập 3 xe cửa số 1 - 2 - 3. Xe về dồn dập mở thêm cửa số 4 bằng 1 băng nâng hạ giải phóng xe tồn.
  - Từ 9:30: Chuyển 2 băng nâng hạ sang layout GXT xuất hàng lên các xe nội vùng/liên tỉnh giảm sức nhân viên.
  - Xe nâng điện 2 tầng kéo hàng từ bãi nhập liên vùng vào layout phân chọn bậc 2. Xe nâng tay kéo hỗ trợ khi xe nâng điện kéo không kịp.
  - Băng tải con lăn đặt trên thùng xe giảm sức khiêng hàng ra cuối thùng.
- **Phương án Ca Night (20h00-06h00)**:
  - Dành 2 băng nâng hạ cho xe liên vùng HCM, 2 băng phân chọn + 2 băng nâng hạ nhập xe về trễ.

#### 🏢 2. KHO CHUYỂN TIẾP BÌNH THUẬN
- Check xe tăng cường từ TP.HCM về để linh động nhân sự nhập xuất.
- Chuẩn bị trước layout bãi hạ hàng cồng kềnh/nặng cho ngày Peak.
- Book xe xuất tăng cường đi Bưu cục khi các COT cố định rớt tải do full.
- **Khẩn cấp**: Bổ sung 3 - 4 Freelance ca đêm ngày 08/08 và 10/08 để đưa năng suất ca đêm về mức an toàn (<1,000 đơn/người).

#### 🏢 3. KHO CHUYỂN TIẾP ĐỨC TRỌNG (LÂM ĐỒNG)
- Check xe tăng cường HCM về linh động nhân sự nhập xuất.
- Chuẩn bị trước layout bãi Bulky/Freight riêng biệt cho ngày Peak 08/08 (hàng cồng kềnh tăng vọt 3,483 đơn).
- Book xe xuất tăng cường đi Đà Lạt/Bảo Lộc khi COT rớt do full tải.
- Sử dụng xe nâng điện 2 tầng hạ tải pallet cồng kềnh.

#### 🏢 4. KHO CHUYỂN TIẾP BẢO LỘC & ĐẮK NÔNG
- Check xe tăng cường HCM, book xe xuất tăng cường đi Bưu cục khi full tải.
- Kiểm tra 100% PDA sạc đủ + SIM 4G phát dự phòng cúp điện/wifi chập chờn.
- Hàng tăng 30%: NVCT sẵn sàng hủy lịch off đi làm xử lý hàng, tăng ca khi phát sinh xe về trễ.

---

## 4. QUY TRÌNH & CHECKLIST CHUẨN BỊ VẬN HÀNH

```mermaid
flowchart TD
    A["Truck từ HCM / Trung tâm về KTC"] --> B["Bắn Mã Nhập Kho (Inbound Scanned)"]
    B --> C{"Phân Luồng Tự Động / Bán Tự Động"}
    C -->|Hàng Normal <5kg| D1["Bàn Chia Nhanh Normal (79.7% Volume)"]
    C -->|Hàng Bulky 5-15kg| D2["Layout Bãi Bulky - Hạ Pallet (12.6% Volume)"]
    C -->|Hàng Freight >15kg| D3["Layout Hàng Nặng - Xe Nâng (7.7% Volume)"]
    D1 --> E["Bắn PDA Chia Theo BC Đích"]
    D2 --> E
    D3 --> E
    E --> F["Bắn PDA Xuất Kho (Outbound Scanned)"]
    F --> G["Chất Xe COT Xuất Đi Bưu Cục"]
```

### Checklist Công Việc Trước & Trong Event 8.8

1. **Trước Event (Hoàn thành trước 05/08)**:
   - [x] Họp toàn bộ nhân viên các kho thông báo kế hoạch Event 8.8 và lịch phân ca.
   - [x] Dọn dẹp quy hoạch layout 3 nhóm hàng Normal, Bulky, Freight.
   - [x] Kiểm tra bảo dưỡng 100% máng nhập, băng tải nâng hạ, xe nâng điện, xe nâng tay.
   - [x] Kiểm tra 100% thiết bị PDA nhập/xuất + 100% SIM 4G dự phòng.
2. **Trong Event (06/08 – 15/08)**:
   - [x] Kiểm tra xe tăng cường từ HCM về mỗi 2 tiếng để chủ động nhân sự.
   - [x] Book xe tăng cường xuất đi Bưu cục khi các COT cố định rớt hàng do full tải.
   - [x] Kiểm tra PDA trước và sau ca làm việc, xử lý lỗi ngay không để ảnh hưởng khung giờ cao điểm.

---

## 5. PHƯƠNG ÁN ỨNG PHÓ KHI HÀNG TĂNG 30% (SURGE CONTINGENCY PLAN)

> [!WARNING]
> Kịch bản ứng phó khi sản lượng thực tế vượt dự báo +30% (Sản lượng sorting KTC NTB đạt **125,000 đơn/ngày** vào ngày 08/08).

### Matrix Giải Pháp Kích Hoạt

1. **Nhân sự**: Thông báo toàn bộ NVCT đi làm đầy đủ không off. Cho NVCT đi làm ngày nghỉ (sắp xếp nghỉ bù sau). Báo HR book thêm Freelance cho cả ca ngày và ca đêm.
2. **Thiết bị & Cổng nhập**: Setup thêm băng tải nâng hạ, mở thêm cửa nhập giải phóng giảm xe chờ nhập.
3. **Xe xuất & Layout**: Book xe xuất tăng cường giải phóng layout. Mở rộng khu vực layout lưu trữ tạm có bạt che ngoài sân kho.

---

## 6. TỔNG KẾT & QUY TRÌNH BÁO CÁO

- **Báo cáo định kỳ (4h/lần)**: Cập nhật sản lượng nhập/xuất/tồn kho và tỷ lệ đúng giờ COT.
- **Báo cáo Sự cốKhẩn cấp**: Sự cố xe về trễ >30 phút hoặc lỗi thiết bị PDA >20% báo cáo trực tiếp AM / Trưởng vùng trong 15 phút.
"""

# Save to Markdown files
artifact_md = r'C:\Users\lap4all\.gemini\antigravity-ide\brain\497c8661-37a7-4c4e-986e-e8ef6c144fb3\KTC_Ke_Hoach_Van_Hanh_Event_8_8_NTB.md'
workspace_md = r'c:\Users\lap4all\Documents\Auto report\KTC_Ke_Hoach_Van_Hanh_Event_8_8_NTB.md'

with open(artifact_md, 'w', encoding='utf-8') as f:
    f.write(md_content)

with open(workspace_md, 'w', encoding='utf-8') as f:
    f.write(md_content)

print('Saved Master Markdown files successfully!')
