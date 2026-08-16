import os
from PIL import Image, ImageDraw, ImageFont
import math
import sys

sys.stdout.reconfigure(encoding='utf-8')

artifact_dir = r"C:\Users\lap4all\.gemini\antigravity-ide\brain\5935f00b-cd19-4de8-83bb-03170a8c8606"
src_web_map = os.path.join(artifact_dir, "media__1785930379076.png")

downloads_dir = r"C:\Users\lap4all\Downloads"
out_img_path = os.path.join(downloads_dir, "Da_Lat_Official_Map_Infographic.png")
artifact_out_path = os.path.join(artifact_dir, "da_lat_official_map_infographic.png")

# Canvas Dimensions: 2200 x 1300 (Ultra HD Infographic)
W, H = 2200, 1300
canvas = Image.new("RGB", (W, H), (248, 250, 252))
draw = ImageDraw.Draw(canvas)

# Fonts
try:
    font_title = ImageFont.truetype("arial.ttf", 34)
    font_subtitle = ImageFont.truetype("arial.ttf", 20)
    font_section = ImageFont.truetype("arial.ttf", 19)
    font_bold = ImageFont.truetype("arial.ttf", 17)
    font_body = ImageFont.truetype("arial.ttf", 15)
    font_small = ImageFont.truetype("arial.ttf", 13)
except:
    font_title = font_subtitle = font_section = font_bold = font_body = font_small = ImageFont.load_default()

def draw_star(draw_ctx, cx, cy, radius, fill_color, outline_color=(255, 255, 255, 255)):
    points = []
    for i in range(10):
        r = radius if i % 2 == 0 else radius * 0.45
        angle = i * math.pi / 5 - math.pi / 2
        x = cx + r * math.cos(angle)
        y = cy + r * math.sin(angle)
        points.append((x, y))
    draw_ctx.polygon(points, fill=fill_color, outline=outline_color)

# 1. Header Banner
draw.rectangle([0, 0, W, 100], fill=(15, 23, 42)) # Dark Navy
draw.text((40, 18), "QUY HOẠCH MẠNG LƯỚI BƯU CỤC TP. ĐÀ LẠT - 2026", fill=(255, 255, 255), font=font_title)
draw.text((40, 62), "ĐỀ XUẤT TÁCH MỚI BƯU CỤC (LDO) XUÂN HƯƠNG 2 (TẠI PHƯỜNG 10) TRÊN NỀN WEB ANH THIÊN (quyhoachbuucuc.info)", fill=(203, 213, 225), font=font_subtitle)

# 2. Place Web Map in Center
map_x, map_y = 660, 120
map_w, map_h = 1000, 1070

if os.path.exists(src_web_map):
    map_im = Image.open(src_web_map).convert("RGB")
    map_im = map_im.resize((map_w, map_h), Image.Resampling.LANCZOS)
    canvas.paste(map_im, (map_x, map_y))
    draw.rectangle([map_x, map_y, map_x + map_w, map_y + map_h], outline=(203, 213, 225), width=3)

# Overlay Draw Layer over Map area
map_overlay = Image.new("RGBA", (W, H), (255, 255, 255, 0))
draw_map = ImageDraw.Draw(map_overlay)

# Map Overlay Coordinates:
# Phường 10 (Yellow Area): approx (map_x + 600, map_y + 650)
# Phường 1, 2, 4 (Blue Area): approx (map_x + 420, map_y + 760)

# Highlight Phường 10 with Orange/Yellow Boundary
p10_poly = [
    (map_x + 500, map_y + 550), (map_x + 720, map_y + 500), (map_x + 850, map_y + 620),
    (map_x + 780, map_y + 820), (map_x + 580, map_y + 850), (map_x + 480, map_y + 700)
]
draw_map.polygon(p10_poly, fill=(254, 240, 138, 110), outline=(234, 179, 8, 255))

# Highlight Phường 1, 2, 4 with Blue Boundary
p124_poly = [
    (map_x + 280, map_y + 680), (map_x + 480, map_y + 700), (map_x + 580, map_y + 850),
    (map_x + 520, map_y + 980), (map_x + 320, map_y + 950), (map_x + 250, map_y + 820)
]
draw_map.polygon(p124_poly, fill=(147, 197, 253, 110), outline=(30, 64, 175, 255))

# Place Stars on Map
# BC Xuân Hương 2 (MỚI tại Phường 10) -> (map_x + 620, map_y + 680)
star_x2, star_y2 = map_x + 620, map_y + 680
draw_star(draw_map, star_x2, star_y2, radius=18, fill_color=(234, 179, 8, 255))
draw_map.rectangle([star_x2 - 110, star_y2 + 22, star_x2 + 110, star_y2 + 48], fill=(255, 255, 255, 240), outline=(202, 138, 4, 255))
draw_map.text((star_x2 - 102, star_y2 + 25), "🟄 (2) BC Xuân Hương 2 (MỚI)", fill=(161, 98, 7, 255), font=font_bold)

# BC Xuân Hương 1 (CŨ) -> (map_x + 400, map_y + 800)
star_x1, star_y1 = map_x + 400, map_y + 800
draw_star(draw_map, star_x1, star_y1, radius=18, fill_color=(30, 64, 175, 255))
draw_map.rectangle([star_x1 - 110, star_y1 + 22, star_x1 + 110, star_y1 + 48], fill=(255, 255, 255, 240), outline=(30, 64, 175, 255))
draw_map.text((star_x1 - 102, star_y1 + 25), "★ (7) BC Xuân Hương 1 (CỦ)", fill=(30, 64, 175, 255), font=font_bold)

# Composite Map Overlay
canvas.paste(Image.alpha_composite(Image.new("RGBA", (W, H), (0,0,0,0)), map_overlay).convert("RGB"), (0,0), map_overlay)

# 3. Left Side Panel: Proposal Cards & Detailed Reason Box
left_x = 40
left_w = 590

# Card 1: BC Xuân Hương 1 (CŨ) - Blue Header
draw.rectangle([left_x, 120, left_x + left_w, 165], fill=(30, 64, 175))
draw.text((left_x + 15, 132), "(7) (LDO) XUÂN HƯƠNG - ĐÀ LẠT 1 (BC CŨ)", fill=(255, 255, 255), font=font_section)

draw.rectangle([left_x, 165, left_x + left_w, 320], fill=(255, 255, 255), outline=(203, 213, 225), width=2)
c1_info = [
    "🚚 Sản lượng giao: 1.400 đơn/ngày  |  📦 Sản lượng lấy: 400 đơn/ngày",
    "👥 Định biên nhân sự: 13 NVPTTT + 1 NVXL (Tổng 14 người)",
    "📍 Phạm vi địa bàn phụ trách (3 phường):",
    "   • Phường 1 (400 đơn/ngày)",
    "   • Phường 2 (400 đơn/ngày)",
    "   • Phường 4 (600 đơn/ngày)"
]
cy = 180
for line in c1_info:
    f_w = font_bold if "Sản lượng" in line or "Phạm vi" in line else font_body
    color = (30, 64, 175) if "Phạm vi" in line else (30, 41, 59)
    draw.text((left_x + 15, cy), line, fill=color, font=f_w)
    cy += 23

# Card 2: BC Xuân Hương 2 (MỚI TẠI PHƯỜNG 10) - Yellow/Pink Header
draw.rectangle([left_x, 340, left_x + left_w, 385], fill=(202, 138, 4))
draw.text((left_x + 15, 352), "(2) (LDO) XUÂN HƯƠNG - ĐÀ LẠT 2 (MỚI TẠI PHƯỜNG 10)", fill=(255, 255, 255), font=font_section)

draw.rectangle([left_x, 385, left_x + left_w, 540], fill=(254, 243, 199), outline=(234, 179, 8), width=2)
c2_info = [
    "🚚 Sản lượng giao: 900 đơn/ngày  |  📦 Sản lượng lấy: 150 đơn/ngày",
    "👥 Định biên nhân sự: 8 NVPTTT + 1 NVXL (Mới bổ sung 8 NV)",
    "📍 Phạm vi địa bàn phụ trách (2 phường):",
    "   • Phường 10 (400 đơn/ngày - Đặt bưu cục)",
    "   • Phường 3 (500 đơn/ngày)"
]
cy = 400
for line in c2_info:
    f_w = font_bold if "Sản lượng" in line or "Phạm vi" in line else font_body
    color = (161, 98, 7) if "Phạm vi" in line else (15, 23, 42)
    draw.text((left_x + 15, cy), line, fill=color, font=f_w)
    cy += 23

# Reason & Explanation Box (Bottom Left Box)
reason_y = 560
draw.rectangle([left_x, reason_y, left_x + left_w, reason_y + 45], fill=(15, 23, 42))
draw.text((left_x + 15, reason_y + 12), "❖ LÝ DO VÀ GIẢI THÍCH NGUYÊN NHÂN TÁCH BƯU CỤC", fill=(255, 255, 255), font=font_section)

draw.rectangle([left_x, reason_y + 45, left_x + left_w, reason_y + 630], fill=(255, 255, 255), outline=(203, 213, 225), width=2)
reasons = [
    "1. Sản lượng Phường Xuân Hương cực lớn (hơn 2.850 đơn/ngày).",
    "2. Địa hình bị chia cắt bởi Hồ Xuân Hương, giao thông di chuyển",
    "   vòng quanh hồ tốn thời gian và nguy hiểm mùa mưa bão.",
    "3. Đây là khu vực trung tâm TP. Đà Lạt có mật độ dân cư,",
    "   shop online và sản lượng giao nhận vô cùng lớn.",
    "4. Thời tiết ở Đà Lạt những tháng cuối năm thường mưa và lạnh,",
    "   giao hàng đường xa dễ dẫn đến trễ KPI và thiếu hụt nhân sự.",
    "5. Việc giữ 01 bưu cục cũ sẽ gây quá tải trầm trọng cho kho bãi",
    "   và shipper. Tách BC Xuân Hương 2 tại Phường 10 giúp cân đối",
    "   sản lượng, giảm bán kính di chuyển và tối ưu năng suất."
]
cy = reason_y + 60
for r in reasons:
    draw.text((left_x + 15, cy), r, fill=(30, 41, 59), font=font_body)
    cy += 25

# 4. Right Side Panel: Other Hubs Summary in TP. Đà Lạt
right_x = 1680
right_w = 480

draw.rectangle([right_x, 120, right_x + right_w, 165], fill=(15, 23, 42))
draw.text((right_x + 15, 132), "TỔNG QUAN CÁC BC TP. ĐÀ LẠT 2026", fill=(255, 255, 255), font=font_section)

hubs_right = [
    ("(1) (LDO) Lâm Viên 1", "Giao: 1.450 đơn | Lấy: 450 đơn | 9 NV", "Phường 9 (600đ), Phường 12 (400đ)", (254, 243, 199)),
    ("(6) (LDO) Lâm Viên 2", "Giao: 1.000 đơn | Lấy: 100 đơn | 6 NV", "Phường 8 (900đ)", (241, 245, 249)),
    ("(5) (LDO) Lang Biang 1", "Giao: 900 đơn | Lấy: 100 đơn | 7 NV", "Phường 7 (300đ), Lạc Dương, Đa Nhim", (236, 253, 245)),
    ("(4) (LDO) Lang Biang 2", "Giao: 1.500 đơn | Lấy: 500 đơn | 10 NV", "Phường 5 (400đ), Phường 6 (500đ), Tà Nung", (239, 246, 255)),
    ("(3) (LDO) Xuân Trường", "Giao: 900 đơn | Lấy: 150 đơn | 7 NV", "Phường 11 (300đ), Xuân Trường, Trạm Hành", (243, 232, 255))
]

ry = 175
for h_name, h_stats, h_cov, bg_c in hubs_right:
    draw.rectangle([right_x, ry, right_x + right_w, ry + 120], fill=bg_c, outline=(203, 213, 225), width=1)
    draw.text((right_x + 15, ry + 10), h_name, fill=(15, 23, 42), font=font_bold)
    draw.text((right_x + 15, ry + 38), h_stats, fill=(30, 41, 59), font=font_body)
    draw.text((right_x + 15, ry + 68), f"Địa bàn: {h_cov}", fill=(71, 85, 105), font=font_small)
    ry += 130

# Legend Box on Right
draw.rectangle([right_x, ry + 20, right_x + right_w, ry + 160], fill=(255, 255, 255), outline=(203, 213, 225), width=2)
draw.text((right_x + 15, ry + 30), "CHÚ GIẢI KÝ HIỆU BẢN ĐỒ", fill=(15, 23, 42), font=font_bold)
draw_star(draw, right_x + 30, ry + 75, radius=10, fill_color=(234, 179, 8, 255))
draw.text((right_x + 50, ry + 65), "🟄 Bưu cục tách mới (LDO Xuân Hương 2)", fill=(161, 98, 7), font=font_body)

draw_star(draw, right_x + 30, ry + 115, radius=10, fill_color=(30, 64, 175, 255))
draw.text((right_x + 50, ry + 105), "★ Bưu cục giữ nguyên (LDO Xuân Hương 1)", fill=(30, 64, 175), font=font_body)

# 5. Bottom Banner Summary Bar
draw.rectangle([0, H - 90, W, H], fill=(15, 23, 42))
draw.text((60, H - 62), "🚚 Tổng sản lượng giao TP. Đà Lạt: 8.050 đơn/ngày", fill=(255, 255, 255), font=font_section)
draw.text((750, H - 62), "📦 Tổng sản lượng lấy: 2.100 đơn/ngày", fill=(255, 255, 255), font=font_section)
draw.text((1350, H - 62), "👥 Tổng nhân sự toàn TP: 67 NVPTTT + 7 NVXL (+8 NV mới)", fill=(255, 255, 255), font=font_section)

# Save image
canvas.save(out_img_path)
canvas.save(artifact_out_path)
print(f"Successfully generated Da Lat Official Infographic at:\n  - {out_img_path}\n  - {artifact_out_path}")
