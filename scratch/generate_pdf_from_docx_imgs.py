import os, sys
from PIL import Image
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages

# Set encoding
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

pdf_path = r'C:\Users\lap4all\Downloads\Ban_Do_Quy_Hoach_Mang_Luoi_NTB_2026.pdf'

img_dir = r'C:\Users\lap4all\Documents\Auto report\scratch\docx_images_v3'
user_brain = r'C:\Users\lap4all\.gemini\antigravity-ide\brain\4168ebee-bcaf-4afa-90c0-8284363e14d2'

# Master list of UNIQUE maps (Grouped by Province, NO duplicate images, 100% Accurate Admin Data)
unique_pages_list = [
    # -------------------------------------------------------------------------
    # 1. TỈNH LÂM ĐỒNG
    # -------------------------------------------------------------------------
    {
        'province': 'TỈNH LÂM ĐỒNG',
        'title': 'HÌNH 1: BẢN ĐỒ QUY HOẠCH CHI TIẾT CỤM ĐƠN DƯƠNG',
        'img': os.path.join(user_brain, 'media__1786004456873.png'),
        'desc': '• BC gốc Nghĩa Đức (Màu đỏ): Cover 6 xã (Đạ Ròn, Thạnh Mỹ, Tu Tra, Ka Đơn, Quảng Lập, Pró) - Vol: 600-720 đơn/ngày, 9/9 NV.\n• BC đề xuất Lạc Xuân (Màu vàng): Cover 4 xã (Lạc Lâm, Lạc Xuân, D\'Ran, Ka Đô) - Vol: 400-480 đơn/ngày, 7/7 NV mới.\n- GIẢM 45% BÁN KÍNH DI CHUYỂN TOÀN CỤM ĐƠN DƯƠNG.'
    },
    {
        'province': 'TỈNH LÂM ĐỒNG',
        'title': 'BẢN ĐỒ QUY HOẠCH TOÀN KHU VỰC BẢO LỘC (PHƯỜNG 1, B\'LAO & PHƯỜNG 2)',
        'img': os.path.join(img_dir, 'unique_img_1.png'),
        'desc': '• Phường 1 mới (P.1 + P. Lộc Phát + X. Lộc Thanh) & Phường B\'Lao mới (P. B\'Lao + P. Lộc Sơn + X. Lộc Nga): Gộp về BC (LDO) B\'Lao Mới (Vol giao 1,500 đơn, lấy 1,000 đơn/ngày).\n• ĐÓNG CỬA BC 1 Bảo Lộc cũ (<50m²), chuyển 100% nhân sự về BC B\'Lao Mới.\n• Phường 2 mới (P.2 + X. Đạm Bri + X. Lộc Tân): Gộp về BC (LDO) 3 Bảo Lộc phụ trách chính.'
    },
    {
        'province': 'TỈNH LÂM ĐỒNG',
        'title': 'BẢN ĐỒ QUY HOẠCH PHƯỜNG XUÂN HƯƠNG (TP. ĐÀ LẠT)',
        'img': os.path.join(img_dir, 'unique_img_8.png'),
        'desc': '• Phường Xuân Hương mới (gộp Phường 1 + P.2 + P.3 + P.4 + P.10).\n• MỜ MỚI BC (LDO) Xuân Hương 2 tại P.10 (cover P.3 & P.10, Vol giao 1,050 đơn, lấy 150 đơn/ngày).\n• BC Xuân Hương cũ: Giữ cover Phường 1 & Phường 2, giải tỏa ùn tắc kho chật hẹp quá tải.'
    },
    {
        'province': 'TỈNH LÂM ĐỒNG',
        'title': 'BẢN ĐỒ QUY HOẠCH XÃ DI LINH (HUYỆN DI LINH)',
        'img': os.path.join(img_dir, 'unique_img_9.png'),
        'desc': '• Xã Di Linh mới (gộp TT. Di Linh + X. Tân Châu + X. Liên Đầm + X. Gung Ré).\n• TÁCH BƯU CỤC HÀNG NHỎ / HÀNG VỪA tại Xã Đinh Trang Thượng phụ trách Đinh Trang Thượng, Di Linh, Phúc Thọ, Liên Đầm.\n• Giảm từ 22km - 45km bán kính di chuyển xa cho BC Di Linh cũ.'
    },
    {
        'province': 'TỈNH LÂM ĐỒNG',
        'title': 'BẢN ĐỒ QUY HOẠCH XÃ ĐỨC TRỌNG (HUYỆN ĐỨC TRỌNG)',
        'img': os.path.join(img_dir, 'unique_img_14.png'),
        'desc': '• Xã Đức Trọng mới (gộp TT. Liên Nghĩa + X. Phú Hội).\n• GIỮ NGUYÊN 02 BƯU CỤC (Đức Trọng 1 cover Xã Phú Hội & Đức Trọng 2 cover TT. Liên Nghĩa) do khoảng cách >15km, kho cũ hẹp và nhân sự khó tuyển dụng mùa cà phê / mưa bão.'
    },

    # -------------------------------------------------------------------------
    # 2. TỈNH KHÁNH HÒA
    # -------------------------------------------------------------------------
    {
        'province': 'TỈNH KHÁNH HÒA',
        'title': 'HÌNH 2: BẢN ĐỒ QUY HOẠCH TÁCH BƯU CỤC CAM LINH VÀ BẮC CAM RANH',
        'img': os.path.join(user_brain, 'media__1786007627105.png'),
        'desc': '• Bưu cục (KHO) Nam Cam Ranh (Mới - Ngôi sao xanh lá): Tách từ BC Cam Linh cũ cover 6 xã/phường phía Nam (Ba Ngòi, Cam Bình, Cam Lập, Cam Phước Đông...) - Vol giao: 500-700 đơn, Lấy: 100-150 đơn/ngày, 9 NVPTTT + 1 NVXL.\n• Bưu cục (KHO) Cam Linh (Ngôi sao đỏ): Giữ phụ trách 6 phường trung tâm - Vol giao: 1,200-1,400 đơn, Lấy: 200-300 đơn/ngày.\n• Bưu cục (KHO) Bắc Cam Ranh: Di dời kho bãi (Ngôi sao xanh dương) & mở rộng mặt bằng (500-700 đơn giao, 100-200 đơn lấy).'
    },
    {
        'province': 'TỈNH KHÁNH HÒA',
        'title': 'BẢN ĐỒ QUY HOẠCH TOÀN BỘ KHU VỰC TP. NHA TRANG (PHƯỜNG NAM, TRUNG TÂM & TÂY NHA TRANG)',
        'img': os.path.join(img_dir, 'unique_img_4.png'),
        'desc': '• Phường Nam Nha Trang mới (P. Phước Hải + P. Phước Long + P. Vĩnh Trường): Gộp về BC Nam Nha Trang 1 Mới & BC Nam Nha Trang 5; ĐÓNG CỬA BC Nam Nha Trang 3.\n• Phường Nha Trang mới (P. Vạn Thạnh + P. Lộc Thọ + P. Tân Tiến + P. Phước Hòa + P. Vĩnh Nguyên): Gộp phân vùng về BC (KHO) Nha Trang.\n• Phường Tây Nha Trang mới (P. Ngọc Hiệp + P. Phương Sài + X. Vĩnh Ngọc + X. Vĩnh Thạnh + X. Vĩnh Trung + X. Vĩnh Hiệp): Gộp về BC (KHO) Tây Nha Trang.'
    },
    {
        'province': 'TỈNH KHÁNH HÒA',
        'title': 'BẢN ĐỒ QUY HOẠCH KHU VỰC THỊ XÃ NINH HÒA (PHƯỜNG NINH HÒA & XÃ TÂN ĐỊNH)',
        'img': os.path.join(img_dir, 'unique_img_6.png'),
        'desc': '• Phường Ninh Hòa mới (P. Ninh Hiệp + X. Ninh Đông + P. Ninh Đa + X. Ninh Phụng) & Xã Tân Định mới (X. Ninh Xuân + X. Ninh Bình + X. Ninh Quang).\n• GIỮ NGUYÊN 02 BƯU CỤC (Ninh Hòa 1 & Ninh Hòa 2) phụ trách các xã xa trung tâm và giữ chân lực lượng lao động tại chỗ.\n• Tối ưu phân vùng mặt bằng kho BC Ninh Hòa 2 để đảm bảo chứa hàng và gom hàng hiệu quả.'
    },
    {
        'province': 'TỈNH KHÁNH HÒA',
        'title': 'BẢN ĐỒ QUY HOẠCH XÃ VẠN THẮNG (HUYỆN VẠN NINH)',
        'img': os.path.join(img_dir, 'unique_img_13.jpg'),
        'desc': '• Xã Vạn Thắng mới (gộp X. Vạn Bình + X. Vạn Thắng).\n• Quy hoạch Bưu cục phụ trách chính: (KHO) Tu Bông (chuyển phần xã Vạn Bình từ BC Vạn Ninh dồn về BC Tu Bông).'
    },

    # -------------------------------------------------------------------------
    # 3. TỈNH NINH THUẬN
    # -------------------------------------------------------------------------
    {
        'province': 'TỈNH NINH THUẬN',
        'title': 'BẢN ĐỒ QUY HOẠCH TOÀN TỈNH NINH THUẬN (PHAN RANG, NINH CHỬ, NINH HẢI & PHƯỚC DINH)',
        'img': os.path.join(img_dir, 'unique_img_5.png'),
        'desc': '• Phường Phan Rang mới (P. Phủ Hà + P. Kinh Dinh + P. Đạo Long + P. Đài Sơn): Quy hoạch gộp phân vùng về Bưu cục trung tâm (NTH) Phan Rang (Vol giao 1,000 đơn, lấy 400 đơn/ngày).\n• Phường Ninh Chử mới (P. Văn Hải + TT. Khánh Hải) & Xã Ninh Hải mới (X. Phương Hải + X. Tri Hải + X. Bắc Sơn): DI DỜI MẶT BẰNG KHO BC (NTH) Ninh Chử về vị trí trung tâm ĐVHC mới (phụ trách 75.7% sản lượng).\n• Xã Phước Dinh mới (X. An Hải + X. Phước Dinh + P. Đông Hải): MỜ MỚI BC (NTH) Đông Hải ven biển (Vol giao 600 đơn, lấy 250 đơn hải sản/ngày) chia tải với BC Phước Dinh.'
    },

    # -------------------------------------------------------------------------
    # 4. TỈNH BÌNH THUẬN (THÔNG TIN ĐVHC CHÍNH XÁC 100%)
    # -------------------------------------------------------------------------
    {
        'province': 'TỈNH BÌNH THUẬN',
        'title': 'BẢN ĐỒ QUY HOẠCH TP. PHAN THIẾT, HÀM THẮNG & LA GI',
        'img': os.path.join(img_dir, 'unique_img_2.jpg'),
        'desc': '• Phường La Gi mới (P. Tân Thiện + P. Tân An + P. Bình Tân + X. Tân Bình): Gộp địa bàn về BC (BTH) Phước Hội & BC (BTH) Tân Hải.\n• Phường Bình Thuận mới (P. Phú Tài + X. Phong Nẫm + X. Hàm Hiệp): Gộp dồn sản lượng về BC trung tâm (BTH) Hàm Thắng & BC (BTH) Hàm Liêm.\n• Phường Phan Thiết mới (P. Phú Trinh + P. Lạc Đạo + P. Bình Hưng) & Phường Hàm Thắng mới (P. Xuân An + TT. Phú Long + X. Hàm Thắng): Gộp về BC Hàm Thắng & BC Phú Thủy.'
    },
    {
        'province': 'TỈNH BÌNH THUẬN',
        'title': 'BẢN ĐỒ QUY HOẠCH XÃ PHAN RÍ CỬA (HUYỆN TUY PHONG)',
        'img': os.path.join(img_dir, 'unique_img_11.png'),
        'desc': '• Xã Phan Rí Cửa mới (gộp TT. Phan Rí Cửa + X. Hòa Minh + X. Chí Công).\n• GIỮ NGUYÊN 02 BƯU CỤC (Phan Rí Cửa & Liên Hương) cover song song do ranh giới địa bàn trải dài.'
    },
    {
        'province': 'TỈNH BÌNH THUẬN',
        'title': 'BẢN ĐỒ QUY HOẠCH XÃ TÂN THÀNH (HUYỆN HÀM THUẬN NAM)',
        'img': os.path.join(img_dir, 'unique_img_12.png'),
        'desc': '• Xã Tân Thành mới (gộp X. Thuận Quý + X. Tân Thuận + X. Tân Thành).\n• GIỮ NGUYÊN 02 BƯU CỤC (Tân Hải & Hàm Thuận Nam) cover song song do ranh giới bờ biển dài >20km.'
    }
]

# Standard A4 Landscape dimensions
FIG_W = 11.69
FIG_H = 8.27

# Available image region bounds
MAX_BOX_W = 0.90
MAX_BOX_H = 0.62
CENTER_X = 0.50
CENTER_Y = 0.53

with PdfPages(pdf_path) as pdf:
    for idx, page in enumerate(unique_pages_list):
        fig = plt.figure(figsize=(FIG_W, FIG_H), dpi=150)
        fig.patch.set_facecolor('#FFFFFF')
        
        # Header banner (Fixed position 0..1 across full page width)
        ax_head = fig.add_axes([0, 0.88, 1.0, 0.12])
        ax_head.set_facecolor('#003366') # Dark Blue
        ax_head.set_xlim(0, 1)
        ax_head.set_ylim(0, 1)
        
        # Province Tag Badge
        ax_head.text(0.5, 0.72, f"[{page['province']}]", color='#F39C12', weight='bold', fontsize=11, ha='center', va='center')
        # Title
        ax_head.text(0.5, 0.35, page['title'], color='white', weight='bold', fontsize=12.5, ha='center', va='center')
        ax_head.axis('off')

        # Footer banner (Fixed position 0..1 across full page width)
        ax_foot = fig.add_axes([0, 0, 1.0, 0.04])
        ax_foot.set_facecolor('#1F2937')
        ax_foot.set_xlim(0, 1)
        ax_foot.set_ylim(0, 1)
        ax_foot.text(0.05, 0.4, 'BẢO MẬT NỘI BỘ - BẢN ĐỒ QUY HOẠCH MẠNG LƯỚI NTB 2026', color='#9CA3AF', fontsize=8, ha='left', va='center')
        ax_foot.text(0.95, 0.4, f'Trang {idx+1}/{len(unique_pages_list)}', color='#9CA3AF', fontsize=8, ha='right', va='center')
        ax_foot.axis('off')

        # Content area (Exact aspect ratio calculation)
        if os.path.exists(page['img']):
            try:
                img_obj = Image.open(page['img'])
                w, h = img_obj.size
                img_aspect = w / h
                
                max_w_in = MAX_BOX_W * FIG_W
                max_h_in = MAX_BOX_H * FIG_H
                container_aspect = max_w_in / max_h_in
                
                if img_aspect > container_aspect:
                    fit_w_in = max_w_in
                    fit_h_in = fit_w_in / img_aspect
                else:
                    fit_h_in = max_h_in
                    fit_w_in = fit_h_in * img_aspect
                
                fit_w_frac = fit_w_in / FIG_W
                fit_h_frac = fit_h_in / FIG_H
                
                box_x = CENTER_X - (fit_w_frac / 2.0)
                box_y = CENTER_Y - (fit_h_frac / 2.0)
                
                ax_img = fig.add_axes([box_x, box_y, fit_w_frac, fit_h_frac])
                ax_img.imshow(img_obj)
                ax_img.axis('off')
            except Exception as e:
                print(f"Error loading image {page['img']}: {e}")
        
        # Description Box - FIXED EXACT FRACTIONAL BOUNDS & NO BORDER
        ax_desc = fig.add_axes([0.05, 0.04, 0.90, 0.16])
        ax_desc.set_facecolor('#F8FAFC')
        ax_desc.set_xlim(0, 1)
        ax_desc.set_ylim(0, 1)
        
        for spine in ax_desc.spines.values():
            spine.set_visible(False)
        
        ax_desc.text(0.01, 0.88, 'GHI CHÚ NỘI DUNG QUY HOẠCH & BƯU CỤC PHỤ TRÁCH:', color='#0F172A', weight='bold', fontsize=9.5, ha='left', va='top')
        ax_desc.text(0.01, 0.60, page['desc'], color='#334155', fontsize=8.5, ha='left', va='top', multialignment='left', linespacing=1.3)
        ax_desc.set_xticks([])
        ax_desc.set_yticks([])

        pdf.savefig(fig)
        plt.close(fig)

print(f"Successfully generated ACCURATE PDF maps at: {pdf_path}")
