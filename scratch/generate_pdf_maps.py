import os, sys
from PIL import Image
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages

# Set encoding
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

# Output PDF path
pdf_path = r'C:\Users\lap4all\Downloads\Ban_Do_Quy_Hoach_Mang_Luoi_NTB_2026.pdf'

# List of pages to include (Title, Image path, Description/Notes)
base_brain = r'C:\Users\lap4all\.gemini\antigravity-ide\brain\5935f00b-cd19-4de8-83bb-03170a8c8606'
user_brain = r'C:\Users\lap4all\.gemini\antigravity-ide\brain\4168ebee-bcaf-4afa-90c0-8284363e14d2'

pages_data = [
    {
        'title': 'BẢN ĐỒ QUY HOẠCH CHI TIẾT CỤM ĐƠN DƯƠNG (TỈNH LÂM ĐỒNG)',
        'subtitle': 'Quy hoạch tách mở Bưu cục (LDO) Lạc Xuân (Màu vàng) chia tải cho Bưu cục gốc Nghĩa Đức (Màu đỏ)',
        'img': os.path.join(user_brain, 'media__1786004456873.png'),
        'desc': '• BC gốc Nghĩa Đức: Cover 6 xã (Đạ Ròn, Thạnh Mỹ, Tu Tra, Ka Đơn, Quảng Lập, Pró) - Vol: 600-720 đơn/ngày, 9/9 NV.\n• BC đề xuất Lạc Xuân (Màu vàng): Cover 4 xã (Lạc Lâm, Lạc Xuân, D\'Ran, Ka Đô) - Vol: 400-480 đơn/ngày, 7/7 NV mới.\n👉 GIẢM 45% BÁN KÍNH DI CHUYỂN TOÀN CỤM ĐƠN DƯƠNG.'
    },
    {
        'title': 'BẢN ĐỒ TỔNG QUAN MẠNG LƯỚI BƯU CỤC VÙNG NTB NĂM 2026',
        'subtitle': 'Rà soát 83 Bưu cục Express tại 5 Tỉnh: Khánh Hòa, Ninh Thuận, Bình Thuận, Lâm Đồng, Đắk Nông',
        'img': os.path.join(base_brain, 'web_maps', '00_ntb_region_map.png'),
        'desc': '• Rà soát 36 ĐVHC mới (sáp nhập từ 114 xã/phường cũ).\n• 34/36 ĐVHC mới (94%) hiện đang bị chia cắt tuyến giữa 2-3 Bưu cục cùng phụ trách.\n• Tổng sản lượng toàn vùng: Giao 24,116 đơn/ngày | Lấy 6,067 đơn/ngày.'
    },
    {
        'title': 'BẢN ĐỒ QUY HOẠCH KHU VỰC BẢO LỘC & BẢO LÂM (LÂM ĐỒNG)',
        'subtitle': 'Mở mới BC B\'Lao Mới quy mô lớn, Đóng cửa BC 1 Bảo Lộc cũ chật hẹp',
        'img': os.path.join(base_brain, 'web_maps', 'map_whatif_bao_loc.png'),
        'desc': '• BC (LDO) B\'Lao Mới: Cover trọn Phường 1 và Phường B\'Lao mới (Vol giao 1,500 đơn, lấy 1,000 đơn/ngày, 15 NVPTTT + 2 NVXL).\n• Gộp kho BC 1 Bảo Lộc cũ (<50m²) và điều chuyển toàn bộ nhân sự sang BC B\'Lao Mới.\n• BC 3 Bảo Lộc: Cover Phường 2 Bảo Lộc mới (Phường 2, Đạm Bri, Lộc Tân).'
    },
    {
        'title': 'BẢN ĐỒ QUY HOẠCH KHU VỰC TP. ĐÀ LẠT (LÂM ĐỒNG)',
        'subtitle': 'Mở mới BC (LDO) Xuân Hương - Đà Lạt 2 chia tải cho BC Xuân Hương cũ',
        'img': os.path.join(base_brain, 'web_maps', 'map_whatif_da_lat.png'),
        'desc': '• BC (LDO) Xuân Hương 2 (Mới): Đặt tại Phường 10, cover Phường 3 & Phường 10 (Vol giao 1,050 đơn, lấy 150 đơn/ngày).\n• BC Xuân Hương cũ: Giữ cover Phường 1 & Phường 2, giải tỏa áp lực kho chật hẹp quá tải.'
    },
    {
        'title': 'BẢN ĐỒ QUY HOẠCH KHU VỰC TP. CAM RANH (KHÁNH HÒA)',
        'subtitle': 'Tách mở BC Nam Cam Ranh, Mở rộng kho BC Bắc Cam Ranh',
        'img': os.path.join(base_brain, 'web_maps', 'map_whatif_cam_linh.png'),
        'desc': '• BC Nam Cam Ranh (Mới): Cover 6 xã phía Nam (Ba Ngòi, Cam Bình, Cam Lập, Cam Phước Đông...) - Vol: 500 đơn/ngày.\n• BC Cam Linh cũ: Giữ 6 phường trung tâm (Vol giao 700 đơn, lấy 100 đơn/ngày).\n• BC Bắc Cam Ranh: Mở rộng kho bãi từ 100m² cũ.'
    },
    {
        'title': 'BẢN ĐỒ QUY HOẠCH KHU VỰC TP. NHA TRANG (KHÁNH HÒA)',
        'subtitle': 'Quy hoạch dồn tuyến Nam Nha Trang, Đóng cửa BC Nam Nha Trang 3',
        'img': os.path.join(base_brain, 'web_maps', 'map_whatif_nha_trang.png'),
        'desc': '• BC Nam Nha Trang 1 Mới: Gộp địa bàn và nhân sự từ BC Nam Nha Trang 3 cũ đóng cửa.\n• BC Nha Trang & BC Tây Nha Trang: Gộp phân vùng phụ trách chính Phường Nha Trang và Phường Tây Nha Trang mới.'
    },
    {
        'title': 'BẢN ĐỒ QUY HOẠCH KHU VỰC THỊ XÃ NINH HÒA (KHÁNH HÒA)',
        'subtitle': 'Duy trì chia tuyến song song BC Ninh Hòa 1 & Ninh Hòa 2',
        'img': os.path.join(base_brain, 'web_maps', 'map_whatif_ninh_hoa.png'),
        'desc': '• BC Ninh Hòa 1: Đặt tại các xã xa trung tâm để giữ chân nhân sự tại chỗ và rút ngắn bán kính di chuyển.\n• BC Tu Bông: Quy hoạch dồn toàn bộ sản lượng Xã Vạn Thắng mới (gồm xã Vạn Bình chuyển từ BC Vạn Ninh về).'
    },
    {
        'title': 'BẢN ĐỒ QUY HOẠCH KHU VỰC PHAN RANG & NINH CHỬ (NINH THUẬN)',
        'subtitle': 'Mở mới BC Đông Hải ven biển & Di dời mặt bằng BC Ninh Chử',
        'img': os.path.join(base_brain, 'web_maps', 'map_whatif_phan_rang.png'),
        'desc': '• BC Đông Hải (Mới): Vol giao 600 đơn, lấy 250 đơn/ngày. Tối ưu gom hàng hải sản ven biển.\n• BC Ninh Chử: Di dời kho bãi đến vị trí trung tâm ĐVHC mới sáp nhập.'
    },
    {
        'title': 'BẢN ĐỒ QUY HOẠCH KHU VỰC PHAN THIẾT & NAM THÀNH (BÌNH THUẬN)',
        'subtitle': 'Mở mới BC Nam Thành xóa giao chéo tuyến xa',
        'img': os.path.join(base_brain, 'web_maps', 'map_whatif_phan_thiet.png'),
        'desc': '• BC Nam Thành (Mới): Phụ trách Nam Thành (~250 đơn/ngày) & Nghị Đức (~200 đơn/ngày), giảm di chuyển xa >25km.\n• BC Hàm Thắng: Quy hoạch dồn sản lượng Phường Bình Thuận & Phường Phan Thiết mới.'
    },
    {
        'title': 'BẢN ĐỒ QUY HOẠCH KHU VỰC TP. GIA NGHĨA & ĐẮK NÔNG',
        'subtitle': 'Giữ nguyên bưu cục vận hành song song tại Đắc Sắc, Đức An, Tà Đùng',
        'img': os.path.join(base_brain, 'web_maps', 'map_whatif_gia_nghia.png'),
        'desc': '• Xã Đắc Sắc: Giữ BC Đức Lập & BC Krông Nô (nhân sự Krông Nô không đi theo nếu ép gộp).\n• Xã Đức An: Giữ BC Đức An & BC Trường Xuân (phần Đắk N\'Drung gần kho Trường Xuân hơn).\n• Xã Tà Đùng: Giữ BC Quảng Khê & BC Quảng Sơn (đường đèo dốc >35km).'
    }
]

with PdfPages(pdf_path) as pdf:
    for idx, page in enumerate(pages_data):
        fig = plt.figure(figsize=(11.69, 8.27), dpi=150) # A4 Landscape
        fig.patch.set_facecolor('#FFFFFF')
        
        # Header banner
        ax_head = fig.add_axes([0, 0.90, 1.0, 0.10])
        ax_head.set_facecolor('#003366') # Dark Blue
        ax_head.text(0.5, 0.60, page['title'], color='white', weight='bold', fontsize=14, ha='center', va='center')
        ax_head.text(0.5, 0.25, page['subtitle'], color='#E0E0E0', fontsize=10, ha='center', va='center')
        ax_head.axis('off')

        # Footer banner
        ax_foot = fig.add_axes([0, 0, 1.0, 0.04])
        ax_foot.set_facecolor('#1F2937')
        ax_foot.text(0.05, 0.4, 'BẢO MẬT NỘI BỘ - VÙNG NTB 2026', color='#9CA3AF', fontsize=8, ha='left', va='center')
        ax_foot.text(0.95, 0.4, f'Trang {idx+1}/{len(pages_data)}', color='#9CA3AF', fontsize=8, ha='right', va='center')
        ax_foot.axis('off')

        # Content area
        if os.path.exists(page['img']):
            try:
                img_obj = Image.open(page['img'])
                ax_img = fig.add_axes([0.05, 0.22, 0.90, 0.65])
                ax_img.imshow(img_obj)
                ax_img.axis('off')
            except Exception as e:
                print(f"Error loading image {page['img']}: {e}")
        
        # Description Box at bottom
        ax_desc = fig.add_axes([0.05, 0.05, 0.90, 0.15])
        ax_desc.set_facecolor('#F3F4F6')
        # Add light border
        for spine in ax_desc.spines.values():
            spine.set_color('#D1D5DB')
            spine.set_linewidth(1)
        
        ax_desc.text(0.02, 0.85, '📌 GHI CHÚ QUY HOẠCH & VẬN HÀNH:', color='#1E3A8A', weight='bold', fontsize=9.5, ha='left', va='top')
        ax_desc.text(0.02, 0.60, page['desc'], color='#374151', fontsize=8.5, ha='left', va='top', multialignment='left', linespacing=1.3)
        ax_desc.set_xticks([])
        ax_desc.set_yticks([])

        pdf.savefig(fig, bbox_inches='tight')
        plt.close(fig)

print(f"Successfully generated PDF maps at: {pdf_path}")
