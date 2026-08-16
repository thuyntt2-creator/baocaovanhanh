# -*- coding: utf-8 -*-
import sys, os, docx
from docx.shared import Inches, Pt, RGBColor
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
from PIL import Image, ImageDraw, ImageFont

sys.stdout.reconfigure(encoding='utf-8')

assets_dir = r'c:\Users\lap4all\Documents\Auto report\scratch\ghn_assets'
os.makedirs(assets_dir, exist_ok=True)

# 1. Create Ultra-High Resolution Crisp Banners (Width 2550px = 8.5in at 300DPI)
hdr_path = os.path.join(assets_dir, 'ghn_header_crisp.png')
ftr_path = os.path.join(assets_dir, 'ghn_footer_crisp.png')

# Header Banner (Width 2550px, Height 360px)
img_hdr = Image.new('RGB', (2550, 360), color=(255, 255, 255))
draw_hdr = ImageDraw.Draw(img_hdr)

# Diagonal orange & blue bars flush to right
draw_hdr.polygon([(1000, 50), (2550, 50), (2550, 110), (930, 110)], fill=(250, 100, 0)) # #FA6400
draw_hdr.polygon([(900, 125), (2550, 125), (2550, 230), (810, 230)], fill=(0, 114, 188)) # #0072BC

try:
    font_b_lg = ImageFont.truetype('arialbd.ttf', 88)
    font_s_sm = ImageFont.truetype('ariali.ttf', 38)
except:
    font_b_lg = ImageFont.load_default()
    font_s_sm = ImageFont.load_default()

draw_hdr.text((60, 60), 'GiaoHangNhanh', fill=(250, 100, 0), font=font_b_lg)
draw_hdr.text((64, 175), 'Giao Siêu Nhanh, Giá Siêu Tốt', fill=(0, 114, 188), font=font_s_sm)
img_hdr.save(hdr_path, dpi=(300, 300))
print('Saved Ultra-Crisp Header:', hdr_path)

# Footer Banner (Width 2550px, Height 300px)
img_ftr = Image.new('RGB', (2550, 300), color=(255, 255, 255))
draw_ftr = ImageDraw.Draw(img_ftr)

# Top Orange Angled Section
draw_ftr.polygon([(800, 0), (2550, 0), (2550, 160), (680, 160)], fill=(250, 100, 0))

try:
    font_ftr_title = ImageFont.truetype('arialbd.ttf', 44)
    font_ftr_sub = ImageFont.truetype('arial.ttf', 32)
except:
    font_ftr_title = ImageFont.load_default()
    font_ftr_sub = ImageFont.load_default()

draw_ftr.text((900, 25), 'CÔNG TY CỔ PHẦN DỊCH VỤ GIAO HÀNG NHANH (GHN)', fill=(255, 255, 255), font=font_ftr_title)
draw_ftr.text((980, 95), '📍 Tầng 3, Rivera Park, 7/28 Thành Thái, Phường 14, Quận 10, TP.HCM', fill=(255, 255, 255), font=font_ftr_sub)

# Bottom Blue Bar
draw_ftr.rectangle([(0, 160), (2550, 300)], fill=(0, 114, 188))
draw_ftr.text((90, 205), 'f /GHNexpress', fill=(255, 255, 255), font=font_ftr_sub)
draw_ftr.text((750, 205), 'f /groups/tamsugiaohangcungGHN', fill=(255, 255, 255), font=font_ftr_sub)
draw_ftr.text((1650, 205), '📞 1900 63 66 77', fill=(255, 255, 255), font=font_ftr_sub)
draw_ftr.text((2150, 205), '✉ cskh@ghn.vn', fill=(255, 255, 255), font=font_ftr_sub)
img_ftr.save(ftr_path, dpi=(300, 300))
print('Saved Ultra-Crisp Footer:', ftr_path)

def make_picture_full_bleed_header(run, img_path, width_in_inches=8.5, height_in_inches=1.2, top_offset_in=0.0):
    """Converts inline picture to page-anchored full bleed header image (0 left, 0 top offset)."""
    run.text = ""
    picture = run.add_picture(img_path, width=Inches(width_in_inches), height=Inches(height_in_inches))
    inline = picture._inline
    
    # Convert wp:inline to wp:anchor
    cx = inline.extent.cx
    cy = inline.extent.cy
    docPr = inline.docPr
    graphic = inline.graphic
    
    top_offset_emu = int(top_offset_in * 914400)
    
    anchor_xml = f'''
    <wp:anchor {nsdecls("wp")} {nsdecls("a")} {nsdecls("pic")} {nsdecls("r")}
               distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="251658240"
               behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1">
      <wp:simplePos x="0" y="0"/>
      <wp:positionH relativeFrom="page">
        <wp:posOffset>0</wp:posOffset>
      </wp:positionH>
      <wp:positionV relativeFrom="page">
        <wp:posOffset>{top_offset_emu}</wp:posOffset>
      </wp:positionV>
      <wp:extent cx="{cx}" cy="{cy}"/>
      <wp:effectExtent l="0" t="0" r="0" b="0"/>
      <wp:wrapNone/>
    </wp:anchor>
    '''
    anchor = parse_xml(anchor_xml)
    anchor.append(docPr)
    anchor.append(graphic)
    
    drawing = inline.getparent()
    drawing.replace(inline, anchor)

print('Full Bleed Header function created!')
