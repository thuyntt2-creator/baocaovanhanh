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

# Open image and convert to RGBA
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

# Coordinates on 1024x565 map:
# Thạnh Mỹ (BC Nghĩa Đức): approx (240, 310)
# Xã Lạc Xuân (BC Lạc Xuân): approx (480, 290)

# 1. Highlight Yellow Territory for BC Lạc Xuân (East 4 wards: Lạc Lâm, Lạc Xuân, Ka Đô, D'Ran)
# Draw semi-transparent yellow polygon over East area
yellow_poly = [
    (380, 160), (450, 40), (580, 40), (660, 200),
    (630, 460), (510, 520), (370, 430), (390, 330), (360, 260)
]
draw.polygon(yellow_poly, fill=(254, 240, 138, 120), outline=(234, 179, 8, 255))

# Draw Red Boundary line for BC Nghĩa Đức (West area)
red_poly = [
    (50, 300), (140, 240), (300, 160), (360, 260),
    (390, 330), (370, 430), (220, 530), (70, 420)
]
draw.polygon(red_poly, fill=(239, 68, 68, 60), outline=(220, 38, 38, 255))

# 2. Place Star Markers
# BC Nghĩa Đức (Red Star) at (250, 310)
draw_star(draw, 250, 310, radius=14, fill_color=(220, 38, 38, 255))
draw.rectangle([190, 330, 310, 355], fill=(255, 255, 255, 230), outline=(220, 38, 38, 255))
draw.text((195, 333), "★ BC Nghĩa Đức", fill=(220, 38, 38, 255), font=font_bold)

# BC Lạc Xuân Mở Mới (Yellow Star) at (480, 290)
draw_star(draw, 480, 290, radius=15, fill_color=(202, 138, 4, 255))
draw.rectangle([420, 315, 550, 340], fill=(255, 255, 255, 230), outline=(202, 138, 4, 255))
draw.text((425, 318), "🟄 BC Lạc Xuân (Mới)", fill=(202, 138, 4, 255), font=font_bold)

# 3. Add Callout Annotation Boxes (Matching AM Telegram Graphic Style)

# Callout 1: Red Box (Top-Left)
c1_box = [20, 30, 290, 160]
draw.rectangle(c1_box, fill=(239, 68, 68, 235), outline=(185, 28, 28, 255), width=2)
draw.text((30, 38), "Khu vực BC gốc (Nghĩa Đức)", fill=(255, 255, 255, 255), font=font_bold)
c1_txt = [
    "• Tuyến: Đạ Ròn, Thạnh Mỹ,",
    "  Tu Tra, Ka Đơn, Quảng Lập, Pró.",
    "• Nhân sự: 9/9 NV (Sẵn sàng)",
    "• Sản lượng: 600 - 720 đơn/ngày"
]
cy = 62
for line in c1_txt:
    draw.text((30, cy), line, fill=(255, 255, 255, 255), font=font_small)
    cy += 20

# Callout 2: Yellow Box (Center-East)
c2_box = [310, 30, 620, 160]
draw.rectangle(c2_box, fill=(254, 240, 138, 240), outline=(202, 138, 4, 255), width=2)
draw.text((320, 38), "Khu vực BC đề xuất (Màu vàng)", fill=(15, 23, 42, 255), font=font_bold)
c2_txt = [
    "• Tuyến: Lạc Lâm, Xã Lạc Xuân, D'Ran, Ka Đô.",
    "• Nhân sự: 7/7 NV (Đề xuất mới 7).",
    "• Sản lượng: 400 - 480 đơn/ngày.",
    "• Giảm tải 45% bán kính di chuyển."
]
cy = 62
for line in c2_txt:
    draw.text((320, cy), line, fill=(15, 23, 42, 255), font=font_small)
    cy += 20

# 4. Add Legend Box (Top Right overlay on top of web header)
leg_box = [640, 10, 980, 65]
draw.rectangle(leg_box, fill=(255, 255, 255, 245), outline=(203, 213, 225, 255), width=2)
draw_star(draw, 660, 28, radius=8, fill_color=(202, 138, 4, 255))
draw.text((678, 20), "Bưu cục tách mới (Lạc Xuân)", fill=(15, 23, 42, 255), font=font_small)

draw_star(draw, 660, 48, radius=8, fill_color=(220, 38, 38, 255))
draw.text((678, 40), "Bưu cục hiện hữu (Nghĩa Đức)", fill=(15, 23, 42, 255), font=font_small)

# Composite overlay
final_im = Image.alpha_composite(base_im, overlay)
final_im.convert("RGB").save(out_img_path)
final_im.convert("RGB").save(artifact_out_path)

print(f"Successfully drawn on user's exact web map screenshot!\nSaved at:\n  - {out_img_path}\n  - {artifact_out_path}")
