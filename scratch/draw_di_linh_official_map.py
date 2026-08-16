import os
from PIL import Image, ImageDraw, ImageFont
import math
import sys

sys.stdout.reconfigure(encoding='utf-8')

artifact_dir = r"C:\Users\lap4all\.gemini\antigravity-ide\brain\5935f00b-cd19-4de8-83bb-03170a8c8606"
src_img_path = os.path.join(artifact_dir, "media__1785944103962.png")

downloads_dir = r"C:\Users\lap4all\Downloads"
out_img_path = os.path.join(downloads_dir, "Di_Linh_Official_Web_Drawn_Map.png")
artifact_out_path = os.path.join(artifact_dir, "di_linh_official_web_drawn_map.png")

if not os.path.exists(src_img_path):
    print(f"Source image not found: {src_img_path}")
    sys.exit(1)

base_im = Image.open(src_img_path).convert("RGBA")
W, H = base_im.size
print(f"Source image dimensions: {W} x {H}")

# Create overlay layer
overlay = Image.new("RGBA", (W, H), (255, 255, 255, 0))
draw = ImageDraw.Draw(overlay)

# Fonts
try:
    font_title = ImageFont.truetype("arial.ttf", max(14, int(H * 0.028)))
    font_bold = ImageFont.truetype("arial.ttf", max(12, int(H * 0.024)))
    font_small = ImageFont.truetype("arial.ttf", max(10, int(H * 0.020)))
except:
    font_title = font_bold = font_small = ImageFont.load_default()

def draw_star(draw_ctx, cx, cy, radius, fill_color, outline_color=(255, 255, 255, 255)):
    points = []
    for i in range(10):
        r = radius if i % 2 == 0 else radius * 0.45
        angle = i * math.pi / 5 - math.pi / 2
        x = cx + r * math.cos(angle)
        y = cy + r * math.sin(angle)
        points.append((x, y))
    draw_ctx.polygon(points, fill=fill_color, outline=outline_color)

# Map Coordinates on 1024x558:
# Đinh Trang Thượng: approx (240, 190)
# Di Linh Hub: approx (310, 410)
# Phúc Thọ Lâm Hà: approx (370, 40)
# Liên Đầm: approx (350, 440)

# 1. Highlight Yellow/Orange Territory for BC Hàng Nhỏ / Hàng Vừa (Xã Đinh Trang Thượng & Phúc Thọ)
dtt_poly = [
    (150, 150), (210, 80), (320, 130), (380, 200),
    (310, 280), (210, 290), (160, 240)
]
draw.polygon(dtt_poly, fill=(254, 240, 138, 120), outline=(234, 179, 8, 255))

# Highlight Orange Boundary for BC Di Linh
dilinh_poly = [
    (210, 290), (310, 280), (410, 280), (420, 480),
    (340, 540), (280, 520), (210, 390)
]
draw.polygon(dilinh_poly, fill=(249, 115, 22, 60), outline=(234, 88, 12, 255))

# 2. Place Star Markers
# BC Hàng Nhỏ / Hàng Vừa (Yellow Star) at Đinh Trang Thượng (240, 190)
star_new_x, star_new_y = 240, 190
draw_star(draw, star_new_x, star_new_y, radius=16, fill_color=(202, 138, 4, 255))
draw.rectangle([star_new_x - 110, star_new_y + 20, star_new_x + 110, star_new_y + 44], fill=(255, 255, 255, 240), outline=(202, 138, 4, 255))
draw.text((star_new_x - 102, star_new_y + 23), "🟄 BC Hàng Nhỏ / Vừa (MỚI)", fill=(161, 98, 7, 255), font=font_bold)

# BC Di Linh Hiện Hữu (Red Star) at (310, 410)
star_old_x, star_old_y = 310, 410
draw_star(draw, star_old_x, star_old_y, radius=15, fill_color=(220, 38, 38, 255))
draw.rectangle([star_old_x - 85, star_old_y + 20, star_old_x + 85, star_old_y + 44], fill=(255, 255, 255, 240), outline=(220, 38, 38, 255))
draw.text((star_old_x - 78, star_old_y + 23), "★ BC (LDO) Di Linh", fill=(220, 38, 38, 255), font=font_bold)

# 3. Clean Floating Callout Annotation Cards (Matching AM Graphic Style)

# Card 1: Yellow Callout (Top Left Overlay) - BC Hàng Nhỏ / Hàng Vừa
c1_box = [15, 15, 360, 145]
draw.rectangle(c1_box, fill=(254, 243, 199, 245), outline=(202, 138, 4, 255), width=2)
draw.rectangle([15, 15, 360, 45], fill=(202, 138, 4, 255))
draw.text((25, 22), "TÁCH BC HÀNG NHỎ / HÀNG VỪA (MỚI)", fill=(255, 255, 255, 255), font=font_bold)

c1_txt = [
    "• Đặt tại: Xã Đinh Trang Thượng",
    "• Hàng Nhỏ: 300-350 giao, 10-20 lấy | 4 NVPTTT",
    "  (Cover Đinh Trang Thượng, Di Linh)",
    "• Hàng Vừa: 4 NVPTTT (Cover Đinh Trang",
    "  Thượng, Phúc Thọ Lâm Hà, Di Linh, Liên Đầm)",
    "• Giảm bán kính di chuyển xa 22km - 45km."
]
cy = 50
for line in c1_txt:
    draw.text((25, cy), line, fill=(15, 23, 42, 255), font=font_small)
    cy += 15

# Card 2: Red Callout (Bottom Left Overlay) - BC Di Linh
c2_box = [15, 420, 330, 545]
draw.rectangle(c2_box, fill=(255, 255, 255, 245), outline=(220, 38, 38, 255), width=2)
draw.rectangle([15, 420, 330, 450], fill=(220, 38, 38, 255))
draw.text((25, 426), "BC (LDO) DI LINH (HIỆN HỮU)", fill=(255, 255, 255, 255), font=font_bold)

c2_txt = [
    "• Cầu địa bàn: 452 đơn / 1.166 kg/ngày",
    "• Phụ trách TT. Di Linh, Tân Châu, Gung Rề.",
    "• Trả tuyến Liên Đầm về đúng địa giới.",
    "• Giảm tải đơn xa cho bưu cục gốc."
]
cy = 456
for line in c2_txt:
    draw.text((25, cy), line, fill=(15, 23, 42, 255), font=font_small)
    cy += 17

# 4. Legend Box on Top Right (above web popup)
leg_box = [650, 10, 990, 60]
draw.rectangle(leg_box, fill=(255, 255, 255, 245), outline=(203, 213, 225, 255), width=2)
draw_star(draw, 668, 26, radius=8, fill_color=(202, 138, 4, 255))
draw.text((685, 18), "Bưu cục tách mới (Hàng Nhỏ / Vừa)", fill=(15, 23, 42, 255), font=font_small)

draw_star(draw, 668, 46, radius=8, fill_color=(220, 38, 38, 255))
draw.text((685, 38), "Bưu cục hiện hữu (BC Di Linh)", fill=(15, 23, 42, 255), font=font_small)

# Composite overlay
final_im = Image.alpha_composite(base_im, overlay)
final_im.convert("RGB").save(out_img_path)
final_im.convert("RGB").save(artifact_out_path)

print(f"Successfully drawn Di Linh official web map!\nSaved at:\n  - {out_img_path}\n  - {artifact_out_path}")
