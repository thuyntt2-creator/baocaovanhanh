import os
from PIL import Image, ImageDraw, ImageFont
import math
import sys

sys.stdout.reconfigure(encoding='utf-8')

artifact_dir = r"C:\Users\lap4all\.gemini\antigravity-ide\brain\5935f00b-cd19-4de8-83bb-03170a8c8606"
src_img_path = os.path.join(artifact_dir, "media__1785929590215.png")

downloads_dir = r"C:\Users\lap4all\Downloads"
out_img_path = os.path.join(downloads_dir, "Don_Duong_Official_Web_Drawn_Map.png")
artifact_out_path = os.path.join(artifact_dir, "don_duong_official_web_drawn_map.png")

if not os.path.exists(src_img_path):
    print(f"Source image not found: {src_img_path}")
    sys.exit(1)

base_im = Image.open(src_img_path).convert("RGBA")
W, H = base_im.size

overlay = Image.new("RGBA", (W, H), (255, 255, 255, 0))
draw = ImageDraw.Draw(overlay)

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

# 1. Highlight Yellow Territory for BC Lạc Xuân (East 4 wards: Lạc Lâm, Lạc Xuân, Ka Đô, D'Ran)
# Draw semi-transparent yellow polygon over East area
yellow_poly = [
    (360, 150), (450, 20), (580, 20), (660, 190),
    (630, 460), (510, 520), (370, 430), (380, 310), (350, 250)
]
draw.polygon(yellow_poly, fill=(254, 240, 138, 110), outline=(234, 179, 8, 255))

# Draw Red Boundary line for BC Nghĩa Đức (West area)
red_poly = [
    (50, 300), (140, 240), (290, 150), (350, 250),
    (380, 310), (370, 430), (220, 530), (70, 420)
]
draw.polygon(red_poly, fill=(239, 68, 68, 50), outline=(220, 38, 38, 255))

# 2. Precise Geographical Star Coordinates:
# BC Nghĩa Đức: Located at TT. Thạnh Mỹ (blue circle target icon on web map) -> exact (246, 312)
star1_x, star1_y = 246, 312
draw_star(draw, star1_x, star1_y, radius=14, fill_color=(220, 38, 38, 255))
draw.rectangle([star1_x - 60, star1_y + 18, star1_x + 65, star1_y + 40], fill=(255, 255, 255, 235), outline=(220, 38, 38, 255))
draw.text((star1_x - 55, star1_y + 20), "★ BC Nghĩa Đức", fill=(220, 38, 38, 255), font=font_bold)

# BC Lạc Xuân Mở Mới: Located at Xã Lạc Xuân -> exact (440, 295)
star2_x, star2_y = 440, 295
draw_star(draw, star2_x, star2_y, radius=15, fill_color=(202, 138, 4, 255))
draw.rectangle([star2_x - 65, star2_y + 18, star2_x + 70, star2_y + 40], fill=(255, 255, 255, 235), outline=(202, 138, 4, 255))
draw.text((star2_x - 60, star2_y + 20), "🟄 BC Lạc Xuân (Mới)", fill=(202, 138, 4, 255), font=font_bold)

# 3. Callout Annotation Boxes (Top Left & Center-East)
c1_box = [20, 20, 290, 150]
draw.rectangle(c1_box, fill=(239, 68, 68, 235), outline=(185, 28, 28, 255), width=2)
draw.text((30, 28), "Khu vực BC gốc (Nghĩa Đức)", fill=(255, 255, 255, 255), font=font_bold)
c1_txt = [
    "• Tuyến: Đạ Ròn, Thạnh Mỹ,",
    "  Tu Tra, Ka Đơn, Quảng Lập, Pró.",
    "• Nhân sự: 9/9 NV (Sẵn sàng)",
    "• Sản lượng: 600 - 720 đơn/ngày"
]
cy = 52
for line in c1_txt:
    draw.text((30, cy), line, fill=(255, 255, 255, 255), font=font_small)
    cy += 20

c2_box = [305, 20, 615, 150]
draw.rectangle(c2_box, fill=(254, 240, 138, 240), outline=(202, 138, 4, 255), width=2)
draw.text((315, 28), "Khu vực BC đề xuất (Màu vàng)", fill=(15, 23, 42, 255), font=font_bold)
c2_txt = [
    "• Tuyến: Lạc Lâm, Xã Lạc Xuân, D'Ran, Ka Đô.",
    "• Nhân sự: 7/7 NV (Đề xuất mới 7).",
    "• Sản lượng: 400 - 480 đơn/ngày.",
    "• Giảm tải 45% bán kính di chuyển."
]
cy = 52
for line in c2_txt:
    draw.text((315, cy), line, fill=(15, 23, 42, 255), font=font_small)
    cy += 20

# 4. Legend Box on Top Right Header
leg_box = [630, 10, 970, 65]
draw.rectangle(leg_box, fill=(255, 255, 255, 245), outline=(203, 213, 225, 255), width=2)
draw_star(draw, 650, 28, radius=8, fill_color=(202, 138, 4, 255))
draw.text((668, 20), "Bưu cục tách mới (Lạc Xuân)", fill=(15, 23, 42, 255), font=font_small)

draw_star(draw, 650, 48, radius=8, fill_color=(220, 38, 38, 255))
draw.text((668, 40), "Bưu cục hiện hữu (Nghĩa Đức)", fill=(15, 23, 42, 255), font=font_small)

# Save image
final_im = Image.alpha_composite(base_im, overlay)
final_im.convert("RGB").save(out_img_path)
final_im.convert("RGB").save(artifact_out_path)

print(f"Refined geographical star coordinates!\nSaved at:\n  - {out_img_path}\n  - {artifact_out_path}")
