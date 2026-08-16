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

# Create overlay canvas
overlay = Image.new("RGBA", (W, H), (255, 255, 255, 0))
draw = ImageDraw.Draw(overlay)

# Fonts
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

# 1. Accurately highlight Đinh Trang Thượng (the Pink Polygon at top center, x=200..310, y=140..250)
dtt_poly = [
    (180, 160), (220, 140), (280, 150), (320, 190),
    (310, 240), (270, 260), (210, 250), (180, 210)
]
draw.polygon(dtt_poly, fill=(244, 114, 182, 100), outline=(219, 39, 119, 255), width=2) # Bright pink highlight

# Highlight Di Linh central territory (x=210..370, y=260..480)
dilinh_poly = [
    (210, 260), (270, 260), (310, 240), (360, 280), (380, 360),
    (360, 480), (290, 520), (220, 400)
]
draw.polygon(dilinh_poly, fill=(59, 130, 246, 60), outline=(37, 99, 235, 255), width=2)

# 2. Place Markers (Location Pin / Star)
# Pin 1: Blue Star at Đinh Trang Thượng -> exact (240, 190)
pin1_x, pin1_y = 240, 190
draw_star(draw, pin1_x, pin1_y, radius=14, fill_color=(37, 99, 235, 255))
# Text tag next to pin
draw.rectangle([pin1_x + 18, pin1_y - 12, pin1_x + 165, pin1_y + 12], fill=(255, 255, 255, 240), outline=(37, 99, 235, 255))
draw.text((pin1_x + 24, pin1_y - 9), "📍 BC Đinh Trang Thượng", fill=(37, 99, 235, 255), font=font_bold)

# Pin 2: Red Star at BC Di Linh gốc -> exact (310, 390)
pin2_x, pin2_y = 310, 390
draw_star(draw, pin2_x, pin2_y, radius=14, fill_color=(220, 38, 38, 255))
# Text tag next to pin
draw.rectangle([pin2_x + 18, pin2_y - 12, pin2_x + 145, pin2_y + 12], fill=(255, 255, 255, 240), outline=(220, 38, 38, 255))
draw.text((pin2_x + 24, pin2_y - 9), "📍 (LDO) BC Di Linh", fill=(220, 38, 38, 255), font=font_bold)

# Draw connecting arrows/lines matching HÌNH 2
draw.line([(pin1_x, pin1_y), (180, 70)], fill=(15, 23, 42, 255), width=2)
draw.line([(pin2_x, pin2_y), (480, 80)], fill=(15, 23, 42, 255), width=2)

# 3. Floating Overlay Cards (Top-Left & Top-Center matching HÌNH 2)

# Card 1: Top-Left Orange/Yellow Card (BC Dự Kiến Tách - Đinh Trang Thượng)
c1_box = [15, 15, 260, 130]
draw.rectangle(c1_box, fill=(255, 255, 255, 248), outline=(234, 88, 12, 255), width=2)
draw.rectangle([15, 15, 260, 40], fill=(234, 88, 12, 255))
draw.text((22, 21), "📍 BC DỰ KIẾN TÁCH", fill=(255, 255, 255, 255), font=font_bold)

c1_txt = [
    "• Giao: 400 - 600 đơn/ngày",
    "• Lấy: 30 - 50 đơn/ngày",
    "• Định biên nhân sự: 6/6 NV",
    "• Tuyến cover: Di Linh, Đinh Trang",
    "  Thượng, Phúc Thọ Lâm Hà."
]
cy = 44
for line in c1_txt:
    draw.text((22, cy), line, fill=(15, 23, 42, 255), font=font_small)
    cy += 16

# Card 2: Top-Center Blue Card ((LDO) DI LINH - Kho gốc)
c2_box = [370, 15, 590, 125]
draw.rectangle(c2_box, fill=(255, 255, 255, 248), outline=(37, 99, 235, 255), width=2)
draw.rectangle([370, 15, 590, 40], fill=(37, 99, 235, 255))
draw.text((378, 21), "📍 (LDO) DI LINH (KHO GỐC)", fill=(255, 255, 255, 255), font=font_bold)

c2_txt = [
    "• Giao: 1.300 - 1.600 đơn/ngày",
    "• Lấy: 80 - 150 đơn/ngày",
    "• Định biên nhân sự: 16/16 NV",
    "• Phụ trách cụm trung tâm Di Linh."
]
cy = 44
for line in c2_txt:
    draw.text((378, cy), line, fill=(15, 23, 42, 255), font=font_small)
    cy += 16

# 4. Add Legend & Summary Box Overlay on Top-Right Header (640..720)
leg_box = [600, 15, 725, 140]
draw.rectangle(leg_box, fill=(255, 255, 255, 245), outline=(203, 213, 225, 255), width=2)
draw.text((610, 20), "CHÚ GIẢI 6 XÃ", fill=(15, 23, 42, 255), font=font_bold)

xã_list = [
    ("Di Linh", "450đ · 1.157kg", (34, 197, 94)),
    ("Đinh Trang Thượng", "156đ · 475kg", (236, 72, 153)),
    ("Bảo Thuận", "148đ · 456kg", (249, 115, 22)),
    ("Gia Hiệp", "144đ · 474kg", (59, 130, 246)),
    ("Sơn Điền", "39đ · 144kg", (168, 85, 247)),
    ("Hòa Ninh", "252đ · 654kg", (234, 179, 8))
]
cy = 38
for xa_name, xa_vol, col in xã_list:
    draw.rectangle([610, cy + 3, 618, cy + 11], fill=col)
    draw.text((622, cy), f"{xa_name}: {xa_vol}", fill=(30, 41, 59, 255), font=ImageFont.load_default())
    cy += 15

# Composite overlay
final_im = Image.alpha_composite(base_im, overlay)
final_im.convert("RGB").save(out_img_path)
final_im.convert("RGB").save(artifact_out_path)

print(f"Successfully rendered perfect Di Linh map matching HÌNH 2 & Web Map!\nSaved at:\n  - {out_img_path}\n  - {artifact_out_path}")
