import os
from PIL import Image, ImageDraw, ImageFont
import sys

sys.stdout.reconfigure(encoding='utf-8')

web_dir = r"C:\Users\lap4all\.gemini\antigravity-ide\brain\5935f00b-cd19-4de8-83bb-03170a8c8606\web_maps"
out_dir = os.path.join(web_dir, "custom_drawn")
os.makedirs(out_dir, exist_ok=True)

# Helper to draw legend box with star markers on top-right of map image
def add_map_decorations(img_path, save_path, legend_items, star_markers=None, boundaries=None):
    if not os.path.exists(img_path):
        print(f"File not found: {img_path}")
        return
        
    im = Image.open(img_path).convert("RGBA")
    w, h = im.size
    
    # Overlay draw canvas
    overlay = Image.new("RGBA", (w, h), (255, 255, 255, 0))
    draw = ImageDraw.Draw(overlay)
    
    # Try loading a bold TTF font or default
    try:
        font_title = ImageFont.truetype("arial.ttf", 20)
        font_legend = ImageFont.truetype("arial.ttf", 18)
    except:
        font_title = ImageFont.load_default()
        font_legend = ImageFont.load_default()
        
    # Draw Legend Box on Top Right
    box_w = 260
    box_h = 30 + len(legend_items) * 32
    margin_right = 30
    margin_top = 30
    
    box_x1 = w - margin_right - box_w
    box_y1 = margin_top
    box_x2 = w - margin_right
    box_y2 = margin_top + box_h
    
    # White semi-transparent box with thin border
    draw.rectangle([box_x1, box_y1, box_x2, box_y2], fill=(255, 255, 255, 230), outline=(200, 200, 200, 255), width=2)
    
    # Render legend items
    y_curr = box_y1 + 15
    for icon_type, color, label in legend_items:
        x_icon = box_x1 + 25
        y_icon = y_curr + 10
        
        if icon_type == "star":
            # Draw star icon
            draw.text((x_icon - 8, y_icon - 12), "★", fill=color, font=font_legend)
        elif icon_type == "blue_star":
            draw.text((x_icon - 8, y_icon - 12), "★", fill=(30, 64, 175, 255), font=font_legend)
        elif icon_type == "red_star":
            draw.text((x_icon - 8, y_icon - 12), "★", fill=(220, 38, 38, 255), font=font_legend)
        elif icon_type == "line":
            draw.line([(x_icon - 12, y_icon), (x_icon + 12, y_icon)], fill=color, width=4)
            
        draw.text((box_x1 + 50, y_curr), label, fill=(30, 41, 59, 255), font=font_legend)
        y_curr += 32
        
    # Draw any star markers at specific pixel locations if provided
    if star_markers:
        for x_p, y_p, color, label in star_markers:
            px = int(x_p * w)
            py = int(y_p * h)
            draw.text((px - 12, py - 16), "★", fill=color, font=ImageFont.truetype("arial.ttf", 32) if hasattr(ImageFont, 'truetype') else font_legend)
            if label:
                draw.text((px + 15, py - 10), label, fill=(0, 0, 0, 255), font=font_legend)
                
    # Composite overlay on original image
    out_im = Image.alpha_composite(im, overlay)
    out_im.convert("RGB").save(save_path)
    print(f"Saved custom drawn map: {save_path}")

# Test adding legend decorations to Nha Trang map
test_img = os.path.join(web_dir, "map_whatif_nha_trang.png")
test_out = os.path.join(out_dir, "decorated_nha_trang.png")

add_map_decorations(
    test_img, test_out,
    legend_items=[
        ("blue_star", (30, 64, 175), "Bưu cục mở mới / di dời"),
        ("red_star", (220, 38, 38), "Bưu cục hiện hữu")
    ]
)
