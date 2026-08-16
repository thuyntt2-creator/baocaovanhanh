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

overlay = Image.new("RGBA", (W, H), (255, 255, 255, 0))
draw = ImageDraw.Draw(overlay)

try:
    font_title = ImageFont.truetype("arial.ttf", max(14, int(H * 0.028)))
    font_bold = ImageFont.truetype("arial.ttf", max(11, int(H * 0.023)))
    font_small = ImageFont.truetype("arial.ttf", max(9, int(H * 0.019)))
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

# 1. Highlight Territory Polygons
# Yellow/Gold Territory Polygon over Đinh Trang Thượng (top-left pink/green area)
dtt_poly = [
    (150, 160), (220, 100), (320, 110), (380, 190),
    (360, 270), (260, 290), (160, 250)
]
draw.polygon(dtt_poly, fill=(254, 240, 138, 120), outline=(234, 179, 8, 255), width=3)

# Orange/Red Territory Polygon over Di Linh (center orange/green area)
dilinh_poly = [
    (210, 290), (260, 290), (360, 270), (410, 280), (420, 470),
    (340, 540), (280, 520), (210, 390)
]
draw.polygon(dilinh_poly, fill=(249, 115, 22, 70), outline=(234, 88, 12, 255), width=3)

# 2. Star Markers & Clean Tags (NO square box '□')
# Gold Star at Đinh Trang Thượng (240, 190)
star_gold_x, star_gold_y = 240, 190
draw_star(draw, star_gold_x, star_gold_y, radius=16, fill_color=(202, 138, 4, 255))
draw.rectangle([star_gold_x - 105, star_gold_y + 20, star_gold_x + 105, star_gold_y + 44], fill=(255, 255, 255, 245), outline=(202, 138, 4, 255), width=2)
draw.text((star_gold_x - 98, star_gold_y + 23), "🟄 BC Hàng Nhỏ / Vừa (MỚI)", fill=(161, 98, 7, 255), font=font_bold)

# Red Star at BC Di Linh (310, 390)
star_red_x, star_red_y = 310, 390
draw_star(draw, star_red_x, star_red_y, radius=16, fill_color=(220, 38, 38, 255))
draw.rectangle([star_red_x - 90, star_red_x - 90 + 175, star_red_y + 20, star_red_y + 44], fill=(255, 255, 255, 245), outline=(220, 38, 38, 255), width=2)
draw.text((star_red_x - 82, star_red_y + 23), "★ BC (LDO) Di Linh", fill=(220, 38, 38, 255), font=font_bold)

# 3. Top-Left Gold Banner Overlay Card
c1_box = [15, 15, 375, 145]
draw.rectangle(c1_box, fill=(254, 243, 199, 248), outline=(202, 138, 4, 255), width=2)
draw.rectangle([15, 15, 375, 45], fill=(202, 138, 4, 255))
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

# 4. Bottom-Left Red Banner Overlay Card
c2_box = [15, 420, 345, 545]
draw.rectangle(c2_box, fill=(255, 255, 255, 248), outline=(220, 38, 38, 255), width=2)
draw.rectangle([15, 420, 345, 450], fill=(220, 38, 38, 255))
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

# 5. Top-Right Legend Box
leg_box = [660, 10, 995, 60]
draw.rectangle(leg_box, fill=(255, 255, 255, 248), outline=(203, 213, 225, 255), width=2)
draw_star(draw, 678, 26, radius=8, fill_color=(202, 138, 4, 255))
draw.text((695, 18), "Bưu cục tách mới (Hàng Nhỏ / Vừa)", fill=(15, 23, 42, 255), font=font_small)

draw_star(draw, 678, 46, radius=8, fill_color=(220, 38, 38, 255))
draw.text((695, 38), "Bưu cục hiện hữu (BC Di Linh)", fill=(15, 23, 42, 255), font=font_small)

# Composite overlay
final_im = Image.alpha_composite(base_im, overlay)
final_im.convert("RGB").save(out_img_path)
final_im.convert("RGB").save(artifact_out_path)

print(f"Successfully rendered perfect Di Linh map style v3!\nSaved at:\n  - {out_img_path}\n  - {artifact_out_path}")
