import docx
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

docx_path = r'C:\Users\lap4all\Downloads\NTB - Báo Cáo Tuần (HRBP - ARD) - Tuần 32.docx'
doc = docx.Document(docx_path)

out_dir = r'C:\Users\lap4all\Documents\Auto report\scratch\w32_docx_images'
os.makedirs(out_dir, exist_ok=True)

image_map = {}
for rel in doc.part.rels.values():
    if 'image' in rel.target_ref:
        img_part = rel.target_part
        filename = os.path.basename(rel.target_ref)
        filepath = os.path.join(out_dir, filename)
        with open(filepath, 'wb') as f:
            f.write(img_part.blob)
        image_map[rel.rId] = filename
        print(f"Extracted {filename} ({len(img_part.blob)} bytes)")

print("\n=== PARAGRAPHS AND ASSOCIATED IMAGES ===")
for i, p in enumerate(doc.paragraphs):
    txt = p.text.strip()
    xml = p._element.xml
    r_ids = []
    if 'graphicData' in xml:
        # Find r:embed
        import re
        embeds = re.findall(r'r:embed="([^"]+)"', xml)
        r_ids.extend(embeds)
    if txt or r_ids:
        imgs = [image_map.get(rid, rid) for rid in r_ids]
        print(f"P{i:3d}: '{txt[:100]}' | Images: {imgs}")
