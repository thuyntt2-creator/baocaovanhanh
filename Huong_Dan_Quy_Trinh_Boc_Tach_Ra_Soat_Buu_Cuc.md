# HUỚNG DẪN QUY TRÌNH 4 BƯỚC BÓC TÁCH RÀ SOÁT BƯU CỤC THEO ĐƠN VỊ HÀNH CHÍNH MỚI

---

### 💡 BẢN CHẤT CỦA 2 FILE DỮ LIỆU BAN ĐẦU:
1. **File 1 (`địa chỉ hiện tại.xlsx`):** Là **Bản đồ CŨ** (Thực tế hiện tại: Xã cũ nào do Bưu cục nào giao, Quản lý AM nào phụ trách).
2. **File 2 (`NTB phân tuyến hành chính.xlsx`):** Là **Bản đồ MỚI** (Sắp tới Nhà nước gộp 3-4 xã cũ thành 1 Phường/Xã MỚI).

---

### 📌 QUY TRÌNH 4 BƯỚC BÓC TÁCH RÀ SOÁT THEO LỆNH CÔNG TY:

#### 🔹 **Bước 1: Ghép 2 File lại với nhau (VLOOKUP trong Excel)**
- Dùng mã xã (`ward_code` hoặc `ID GHN xã cũ`) làm chìa khóa liên kết.
- Lấy tên Bưu cục đang chạy thực tế ở **File 1** ghép sang từng dòng xã cũ ở **File 2**.
- 👉 *Mục đích:* Để biết mỗi xã cũ trong sơ đồ sáp nhập hiện tại do Bưu cục nào đảm nhận.

---

#### 🔹 **Bước 2: Gom nhóm (Pivot Table trong Excel) theo "Tên Xã Mới" (Giải quyết Yêu cầu 1)**
- Tạo bảng Pivot Table trong Excel, gom toàn bộ dòng theo cột **`Tên Xã Mới`**.
- Đếm xem mỗi **`Tên Xã Mới`** đang có bao nhiêu Bưu cục nhảy vào giao:
  - Nếu kết quả = **1 Bưu cục** ➔ **ĐÃ CHUẨN** (Phường mới thuộc trọn vẹn 1 Bưu cục).
  - Nếu kết quả **≥ 2 Bưu cục** ➔ **BỊ CHIA CẮT MẢNH** (1 Phường mới nhưng 2-3 Bưu cục cùng giẫm chân lên nhau).
- 👉 *Kết quả bóc tách:* Lọc ra đúng **32 Phường/Xã mới** đang bị chia cắt cho 2-3 Bưu cục.

---

#### 🔹 **Bước 3: Đưa Sản lượng vào để Đánh giá GỘP hay GIỮ (Giải quyết Yêu cầu 2)**
Với 32 Phường/Xã bị chia cắt ở Bước 2, tra cứu Sản lượng đơn/ngày và kg/ngày từ hệ thống:

- **Trường hợp A ➔ Đề xuất GIỮ NGUYÊN 2-3 BƯU CỤC:**
  - **Căn cứ:** Tổng sản lượng Phường mới đó cực kỳ lớn (**> 1.500 đơn/ngày**, ví dụ: *Nam Nha Trang, Nha Trang, Phan Thiết, B'Lao...*).
  - **Lý do giải thích với Sếp:** Nếu ép gộp 1.500 - 2.500 đơn/ngày về 01 Bưu cục duy nhất thì kho đó sẽ **bị vỡ m² (chật cứng kho)**, ùn tắc giờ cao điểm và **không thể tuyển nổi hàng chục shipper mới cùng lúc**.
  - **Giải pháp:** Giữ 2-3 Bưu cục nhưng phân ranh giới tuyến đường theo xã cũ rõ ràng để không chạy cắt mặt nhau.

- **Trường hợp B ➔ Đề xuất GỘP VỀ 01 BƯU CỤC:**
  - **Căn cứ:** Sản lượng vừa phải, hoặc có **01 Bưu cục chính đã gánh > 70% sản lượng** (Bưu cục phụ chỉ dính 10-20% nhỏ lẻ).
  - **Lý do giải thích với Sếp:** Bưu cục chính đã phủ hầu hết phường mới.
  - **Giải pháp:** Cắt phần nhỏ lẻ từ Bưu cục phụ về Bưu cục chính và điều chuyển shipper phụ trách phần đó sang Bưu cục chính quản lý.

---

#### 🔹 **Bước 4: Rà soát Bưu cục quá tải & Tuyến đi chéo xa (Giải quyết Yêu cầu 3)**
- **Tuyến đi chéo xa (Reassign):** So sánh khoảng cách từ Bưu cục hiện tại đến xã so với Bưu cục lân cận. Nếu Bưu cục cũ cách 40km mà Bưu cục bên cạnh chỉ cách 15km ➔ Đề xuất chuyển ngay về Bưu cục gần hơn (Làm xong bước này cho 18 tuyến giúp công ty tiết kiệm **10 - 25km/tuyến**).
- **Bưu cục quá tải m²:** Tìm các kho chật chội có sản lượng cao ➔ Đề xuất mở rộng m² hoặc san sẻ bớt xã cho kho lân cận.
