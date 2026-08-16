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

# Canvas Dimensions: 2040 x 1200 (Map Focused Presentation)
W, H = 2040, 1200
canvas = Image.new("RGB", (W, H), (15, 23, 42)) # Sleek dark background
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
draw.rectangle([0, 0, W, 90], fill=(15, 23, 42))
draw.text((40, 16), "QUY HOẠCH MẠNG LƯỚI BƯU CỤC TP. ĐÀ LẠT - 2026", fill=(255, 255, 255), font=font_title)
draw.text((40, 58), "BẢN ĐỒ PHÂN VÙNG TÁCH MỚI BƯU CỤC (LDO) XUÂN HƯƠNG 2 (TẠI PHƯỜNG 10)", fill=(148, 163, 184), font=font_subtitle)

# 2. Place Web Map as Large Main Focus Canvas (Taking 95% of space)
map_x, map_y = 30, 100
map_w, map_h = 1980, 1010

if os.path.exists(src_web_map):
    map_im = Image.open(src_web_map).convert("RGB")
    map_im = map_im.resize((map_w, map_h), Image.Resampling.LANCZOS)
    canvas.paste(map_im, (map_x, map_y))
    draw.rectangle([map_x, map_y, map_x + map_w, map_y + map_h], outline=(51, 65, 85), width=3)

# Overlay Draw Layer over Map Area
map_overlay = Image.new("RGBA", (W, H), (255, 255, 255, 0))
draw_map = ImageDraw.Draw(map_overlay)

# Map Overlay Polygons & Stars:
# Phường 10 (Yellow Area): approx (map_x + 1180, map_y + 600)
# Phường 1, 2, 4 (Blue Area): approx (map_x + 780, map_y + 720)

p10_poly = [
    (map_x + 950, map_y + 480), (map_x + 1380, map_y + 440), (map_x + 1600, map_y + 580),
    (map_x + 1500, map_y + 780), (map_x + 1100, map_y + 820), (map_x + 900, map_y + 650)
]
draw_map.polygon(p10_poly, fill=(254, 240, 138, 120), outline=(234, 179, 8, 255))

p124_poly = [
    (map_x + 550, map_y + 620), (map_x + 900, map_y + 650), (map_x + 1100, map_y + 820),
    (map_x + 980, map_y + 940), (map_x + 600, map_y + 920), (map_x + 480, map_y + 780)
]
draw_map.polygon(p124_poly, fill=(147, 197, 253, 120), outline=(30, 64, 175, 255))

# Place Stars on Map
# BC Xuân Hương 2 (MỚI tại Phường 10)
star_x2, star_y2 = map_x + 1180, map_y + 620
draw_star(draw_map, star_x2, star_y2, radius=20, fill_color=(234, 179, 8, 255))
draw_map.rectangle([star_x2 - 120, star_y2 + 25, star_x2 + 120, star_y2 + 55], fill=(255, 255, 255, 245), outline=(202, 138, 4, 255))
draw_map.text((star_x2 - 110, star_y2 + 28), "🟄 (2) BC Xuân Hương 2 (MỚI)", fill=(161, 98, 7, 255), font=font_bold)

# BC Xuân Hương 1 (CỦ)
star_x1, star_y1 = map_x + 780, map_y + 750
draw_star(draw_map, star_x1, star_y1, radius=20, fill_color=(30, 64, 175, 255))
draw_map.rectangle([star_x1 - 120, star_y1 + 25, star_x1 + 120, star_y1 + 55], fill=(255, 255, 255, 245), outline=(30, 64, 175, 255))
draw_map.text((star_x1 - 110, star_y1 + 28), "★ (7) BC Xuân Hương 1 (CỦ)", fill=(30, 64, 175, 255), font=font_bold)

# 3. Clean Floating Overlay Cards Over Map (Top Left & Bottom Right)

# Card 1: BC Xuân Hương 1 (Top Left Overlay)
c1_box = [map_x + 30, map_y + 30, map_x + 520, map_y + 220]
draw_map.rectangle(c1_box, fill=(255, 255, 255, 245), outline=(30, 64, 175, 255), width=3)
draw_map.rectangle([map_x + 30, map_y + 30, map_x + 520, map_y + 75], fill=(30, 64, 175, 255))
draw_map.text((map_x + 45, map_y + 42), "(7) (LDO) XUÂN HƯƠNG - ĐÀ LẠT 1 (BC CỦ)", fill=(255, 255, 255, 255), font=font_section)

c1_lines = [
    "🚚 Giao: 1.400 đơn/ngày  |  📦 Lấy: 400 đơn/ngày",
    "👥 Định biên: 13 NVPTTT + 1 NVXL (Tổng 14 người)",
    "📍 Cover 3 phường: Phường 1 (400đ), Phường 2 (400đ),",
    "   Phường 4 (600đ)."
]
cy = map_y + 88
for l in c1_lines:
    draw_map.text((map_x + 45, cy), l, fill=(15, 23, 42, 255), font=font_body)
    cy += 28

# Card 2: BC Xuân Hương 2 MỚI (Bottom Right Overlay)
c2_box = [map_x + 1430, map_y + 680, map_x + 1950, map_y + 870]
draw_map.rectangle(c2_box, fill=(254, 243, 199, 248), outline=(202, 138, 4, 255), width=3)
draw_map.rectangle([map_x + 1430, map_y + 680, map_x + 1950, map_y + 725], fill=(202, 138, 4, 255))
draw_map.text((map_x + 1445, map_y + 692), "(2) (LDO) XUÂN HƯƠNG 2 (MỚI TẠI P.10)", fill=(255, 255, 255, 255), font=font_section)

c2_lines = [
    "🚚 Giao: 900 đơn/ngày  |  📦 Lấy: 150 đơn/ngày",
    "👥 Định biên: 8 NVPTTT + 1 NVXL (Bổ sung 8 NV)",
    "📍 Cover 2 phường: Phường 10 (400đ - Đặt kho),",
    "   Phường 3 (500đ)."
]
cy = map_y + 738
for l in c2_lines:
    draw_map.text((map_x + 1445, cy), l, fill=(15, 23, 42, 255), font=font_body)
    cy += 28

# Legend Box on Top Right Overlay
leg_box = [map_x + 1450, map_y + 30, map_x + 1950, map_y + 115]
draw_map.rectangle(leg_box, fill=(255, 255, 255, 245), outline=(203, 213, 225, 255), width=2)
draw_star(draw_map, map_x + 1475, map_y + 55, radius=10, fill_color=(234, 179, 8, 255))
draw_map.text((map_x + 1495, map_y + 44), "🟄 Bưu cục tách mới (Xuân Hương 2)", fill=(161, 98, 7, 255), font=font_body)

draw_star(draw_map, map_x + 1475, map_y + 90, radius=10, fill_color=(30, 64, 175, 255))
draw_map.text((map_x + 1495, map_y + 80), "★ Bưu cục giữ nguyên (Xuân Hương 1)", fill=(30, 64, 175, 255), font=font_body)

# Composite Map Overlay
canvas.paste(Image.alpha_composite(Image.new("RGBA", (W, H), (0,0,0,0)), map_overlay).convert("RGB"), (0,0), map_overlay)

# 4. Bottom Summary Banner
draw.rectangle([0, H - 80, W, H], fill=(15, 23, 42))
draw.text((60, H - 55), "🚚 Tổng giao Phường Xuân Hương: 2.300 đơn/ngày", fill=(255, 255, 255), font=font_section)
draw.text((720, H - 55), "📦 Tổng lấy: 550 đơn/ngày", fill=(255, 255, 255), font=font_section)
draw.text((1220, H - 55), "👥 Tổng nhân sự 2 bưu cục: 21 NVPTTT + 2 NVXL (Mới +8 NV)", fill=(255, 255, 255), font=font_section)

# Save image
canvas.save(out_img_path)
canvas.save(artifact_out_path)
print(f"Successfully generated Map-Focused Da Lat Infographic at:\n  - {out_img_path}\n  - {artifact_out_path}")
