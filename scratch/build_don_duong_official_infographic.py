import os
from PIL import Image, ImageDraw, ImageFont
import math
import sys

sys.stdout.reconfigure(encoding='utf-8')

artifact_dir = r"C:\Users\lap4all\.gemini\antigravity-ide\brain\5935f00b-cd19-4de8-83bb-03170a8c8606"
web_maps_dir = os.path.join(artifact_dir, "web_maps")
downloads_dir = r"C:\Users\lap4all\Downloads"

out_img_path = os.path.join(downloads_dir, "Don_Duong_Official_Map_Infographic.png")
artifact_img_path = os.path.join(artifact_dir, "don_duong_official_infographic.png")

# Use official map image as base
base_map_path = os.path.join(artifact_dir, "media__1785896953199.png")
if not os.path.exists(base_map_path):
    base_map_path = os.path.join(web_maps_dir, "map_whatif_don_duong.png")

# Canvas Dimensions: 2040 x 1200 (Ultra HD Infographic)
W, H = 2040, 1200
canvas = Image.new("RGB", (W, H), (248, 250, 252)) # light grey-blue background
draw = ImageDraw.Draw(canvas)

# Fonts
try:
    font_title = ImageFont.truetype("arial.ttf", 36)
    font_subtitle = ImageFont.truetype("arial.ttf", 22)
    font_section = ImageFont.truetype("arial.ttf", 20)
    font_body_bold = ImageFont.truetype("arial.ttf", 18)
    font_body = ImageFont.truetype("arial.ttf", 16)
    font_small = ImageFont.truetype("arial.ttf", 14)
except:
    font_title = font_subtitle = font_section = font_body_bold = font_body = font_small = ImageFont.load_default()

# 1. Header Banner
draw.rectangle([0, 0, W, 100], fill=(15, 23, 42)) # Dark Navy Header
draw.text((40, 20), "BẢN ĐỒ PHÂN VÙNG BƯU CỤC ĐƠN DƯƠNG - LÂM ĐỒNG 2026", fill=(255, 255, 255), font=font_title)
draw.text((40, 65), "Trích xuất từ Hệ thống Quy hoạch Bưu cục Anh Thiên (quyhoachbuucuc.info) & Đề xuất Mô hình 2 Bưu cục", fill=(148, 163, 184), font=font_subtitle)

# 2. Place Web Map in Center/Right Area
map_x, map_y = 620, 120
map_w, map_h = 1380, 920

if os.path.exists(base_map_path):
    map_im = Image.open(base_map_path).convert("RGB")
    map_im = map_im.resize((map_w, map_h), Image.Resampling.LANCZOS)
    canvas.paste(map_im, (map_x, map_y))
    # Draw border around map
    draw.rectangle([map_x, map_y, map_x + map_w, map_y + map_h], outline=(203, 213, 225), width=3)

# 3. Left Panel Summary Box (Top Left Table)
tbl_x, tbl_y = 40, 120
tbl_w = 550

# Header card
draw.rectangle([tbl_x, tbl_y, tbl_x + tbl_w, tbl_y + 45], fill=(30, 58, 138)) # Deep Blue
draw.text((tbl_x + 15, tbl_y + 12), "TỔNG HỢP TOÀN KHU VỰC ĐƠN DƯƠNG", fill=(255, 255, 255), font=font_section)

# Table Header
ty = tbl_y + 45
draw.rectangle([tbl_x, ty, tbl_x + tbl_w, ty + 35], fill=(226, 232, 240))
draw.text((tbl_x + 15, ty + 8), "Địa bàn / Bưu cục", fill=(15, 23, 42), font=font_body_bold)
draw.text((tbl_x + 290, ty + 8), "TB giao (đơn/ngày)", fill=(15, 23, 42), font=font_body_bold)
draw.text((tbl_x + 450, ty + 8), "Định biên", fill=(15, 23, 42), font=font_body_bold)

# Table Rows
rows_data = [
    ("Khu vực TT. Thạnh Mỹ & Xã Lạc Xuân", "1.000 - 1.200", "16 người", (241, 245, 249), True),
    ("BC gốc Nghĩa Đức (TT. Thạnh Mỹ)", "600 - 720", "9 người", (255, 255, 255), False),
    ("BC Lạc Xuân Mở Mới (Màu Vàng)", "400 - 480", "7 người", (254, 243, 199), False)
]

ty += 35
for area, vol, staff, bg_color, is_total in rows_data:
    draw.rectangle([tbl_x, ty, tbl_x + tbl_w, ty + 40], fill=bg_color, outline=(203, 213, 225), width=1)
    f_weight = font_body_bold if is_total else font_body
    draw.text((tbl_x + 15, ty + 10), area, fill=(15, 23, 42), font=f_weight)
    draw.text((tbl_x + 310, ty + 10), vol, fill=(15, 23, 42), font=f_weight)
    draw.text((tbl_x + 460, ty + 10), staff, fill=(15, 23, 42), font=f_weight)
    ty += 40

# Legend Details Box below table
leg_y = ty + 20
draw.rectangle([tbl_x, leg_y, tbl_x + tbl_w, leg_y + 190], fill=(255, 255, 255), outline=(203, 213, 225), width=2)
draw.rectangle([tbl_x, leg_y, tbl_x + tbl_w, leg_y + 40], fill=(241, 245, 249))
draw.text((tbl_x + 15, leg_y + 10), "CHÚ GIẢI PHÂN CÔNG QUY HOẠCH", fill=(30, 41, 59), font=font_body_bold)

# Legend item 1: BC Nghĩa Đức
draw.text((tbl_x + 15, leg_y + 50), "★ Bưu cục gốc Nghĩa Đức (Xã Đạ Ròn, TT. Thạnh Mỹ, Tu Tra, Ka Đơn, Quảng Lập, Pró). Nhân sự: 9 người.", fill=(220, 38, 38), font=font_body)
# Legend item 2: BC Lạc Xuân
draw.text((tbl_x + 15, leg_y + 90), "🟄 Bưu cục đề xuất Mở Mới (Xã Lạc Lâm, Xã Lạc Xuân, TT. D'Ran, Ka Đô). Nhân sự: 7 người.", fill=(202, 138, 4), font=font_body)
# Boundary line legend
draw.line([(tbl_x + 15, leg_y + 145), (tbl_x + 65, leg_y + 145)], fill=(220, 38, 38), width=4)
draw.text((tbl_x + 75, leg_y + 135), "Ranh giới BC Nghĩa Đức", fill=(15, 23, 42), font=font_small)

draw.line([(tbl_x + 280, leg_y + 145), (tbl_x + 330, leg_y + 145)], fill=(30, 64, 175), width=4)
draw.text((tbl_x + 340, leg_y + 135), "Ranh giới BC Lạc Xuân", fill=(15, 23, 42), font=font_small)

# 4. Bottom Left Box: New Hub Placement Detail Card
box_y = leg_y + 210
draw.rectangle([tbl_x, box_y, tbl_x + tbl_w, box_y + 240], fill=(15, 23, 42), outline=(30, 58, 138), width=2)
draw.rectangle([tbl_x, box_y, tbl_x + tbl_w, box_y + 45], fill=(30, 58, 138))
draw.text((tbl_x + 15, box_y + 12), "📍 ĐIỂM ĐẶT MỚI BƯU CỤC LẠC XUÂN (ĐỀ XUẤT)", fill=(255, 255, 255), font=font_section)

details_text = [
    "📍 Vị trí đặt kho: Ngôi sao màu vàng (Xã Lạc Xuân)",
    "🌐 Phạm vi dự kiến phục vụ: Các xã phía Đông Bắc (Lạc Lâm, Lạc Xuân, D'Ran, Ka Đô)",
    "👥 Quy mô nhân sự: 7 nhân viên giao nhận + 1 Quản lý",
    "🚚 Sản lượng dự kiến: 400 - 480 đơn/ngày",
    "✅ Tình trạng đề xuất: Đề xuất thành lập mới năm 2026"
]

by = box_y + 55
for line in details_text:
    draw.text((tbl_x + 15, by), line, fill=(241, 245, 249), font=font_body)
    by += 34

# 5. Callout Cards Overlaid on Map Area
# Callout 1: BC Nghĩa Đức (Red Callout)
c1_x, c1_y = map_x + 80, map_y + 160
draw.rectangle([c1_x, c1_y, c1_x + 380, c1_y + 160], fill=(239, 68, 68, 230), outline=(185, 28, 28), width=2)
draw.text((c1_x + 15, c1_y + 10), "Khu vực Bưu cục gốc (Nghĩa Đức)", fill=(255, 255, 255), font=font_body_bold)
c1_desc = [
    "• Tuyến: Đạ Ròn, TT. Thạnh Mỹ, Tu Tra,",
    "  Ka Đơn, Quảng Lập, Pró.",
    "• Nhân sự: 9/9 nhân viên (Sẵn sàng)",
    "• Sản lượng giao: 600 - 720 đơn/ngày"
]
cy = c1_y + 38
for l in c1_desc:
    draw.text((c1_x + 15, cy), l, fill=(255, 255, 255), font=font_small)
    cy += 26

# Callout 2: BC Lạc Xuân (Yellow Callout)
c2_x, c2_y = map_x + 850, map_y + 240
draw.rectangle([c2_x, c2_y, c2_x + 400, c2_y + 170], fill=(234, 179, 8, 230), outline=(161, 98, 7), width=2)
draw.text((c2_x + 15, c2_y + 10), "Khu vực Bưu cục đề xuất (Màu vàng)", fill=(15, 23, 42), font=font_body_bold)
c2_desc = [
    "• Tuyến: Lạc Lâm, Xã Lạc Xuân, D'Ran, Ka Đô.",
    "• Nhân sự: 7/7 nhân viên (Đề xuất mới 7).",
    "• Sản lượng giao: 400 - 480 đơn/ngày",
    "• Tách bớt 45% sản lượng cho kho gốc."
]
cy = c2_y + 38
for l in c2_desc:
    draw.text((c2_x + 15, cy), l, fill=(15, 23, 42), font=font_small)
    cy += 26

# 6. Bottom Banner Bar
draw.rectangle([0, H - 90, W, H], fill=(15, 23, 42))
draw.text((60, H - 65), "🚚 Tổng sản lượng giao: 1.000 - 1.200 đơn/ngày", fill=(255, 255, 255), font=font_section)
draw.text((650, H - 65), "👥 Tổng định biên: 16 người", fill=(255, 255, 255), font=font_section)
draw.text((1150, H - 65), "📊 Hiệu suất trung bình: 68 - 75 đơn/NV/ngày", fill=(255, 255, 255), font=font_section)

# Save image
canvas.save(out_img_path)
canvas.save(artifact_img_path)
print(f"Successfully generated Don Duong Official Infographic at:\n  - {out_img_path}\n  - {artifact_img_path}")
