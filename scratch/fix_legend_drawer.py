import os
from PIL import Image, ImageDraw, ImageFont
import math
import sys

sys.stdout.reconfigure(encoding='utf-8')

web_dir = r"C:\Users\lap4all\.gemini\antigravity-ide\brain\5935f00b-cd19-4de8-83bb-03170a8c8606\web_maps"
custom_dir = os.path.join(web_dir, "custom_drawn")
os.makedirs(custom_dir, exist_ok=True)

def draw_star(draw, cx, cy, radius, fill_color, outline_color=None):
    points = []
    for i in range(10):
        r = radius if i % 2 == 0 else radius * 0.45
        angle = i * math.pi / 5 - math.pi / 2
        x = cx + r * math.cos(angle)
        y = cy + r * math.sin(angle)
        points.append((x, y))
    draw.polygon(points, fill=fill_color, outline=outline_color)

def decorate_map_clean(src_path, dst_name, has_new=True, has_move=True):
    if not os.path.exists(src_path):
        return src_path
    dst_path = os.path.join(custom_dir, dst_name)
    im = Image.open(src_path).convert("RGBA")
    w, h = im.size
    overlay = Image.new("RGBA", (w, h), (255, 255, 255, 0))
    draw = ImageDraw.Draw(overlay)
    
    try:
        font_leg = ImageFont.truetype("arial.ttf", int(h * 0.024))
    except:
        font_leg = ImageFont.load_default()
        
    items = []
    if has_new:
        items.append(((30, 64, 175, 255), "Bưu cục mở mới / tách mới"))
    if has_move:
        items.append(((220, 38, 38, 255), "Bưu cục hiện hữu / di dời"))
        
    box_w = int(w * 0.30)
    box_h = int(h * 0.04) + len(items) * int(h * 0.048)
    box_x1 = w - int(w * 0.03) - box_w
    box_y1 = int(h * 0.03)
    box_x2 = w - int(w * 0.03)
    box_y2 = box_y1 + box_h
    
    # White semi-transparent box
    draw.rectangle([box_x1, box_y1, box_x2, box_y2], fill=(255, 255, 255, 240), outline=(160, 160, 160, 255), width=2)
    
    y_c = box_y1 + int(h * 0.022)
    star_r = int(h * 0.014)
    
    for color, txt in items:
        # Draw solid star polygon
        star_cx = box_x1 + int(w * 0.025)
        star_cy = y_c + int(h * 0.010)
        draw_star(draw, star_cx, star_cy, star_r, fill_color=color, outline_color=(255, 255, 255, 255))
        
        # Text label
        draw.text((box_x1 + int(w * 0.050), y_c), txt, fill=(15, 23, 42, 255), font=font_leg)
        y_c += int(h * 0.048)
        
    out_im = Image.alpha_composite(im, overlay)
    out_im.convert("RGB").save(dst_path)
    return dst_path

# Test decorating
test_src = os.path.join(web_dir, "map_whatif_da_lat.png")
test_dst = decorate_map_clean(test_src, "test_da_lat_clean.png")
print("Saved test clean decorated map:", test_dst)
