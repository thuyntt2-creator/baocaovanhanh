import docx
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

docx_path = r'C:\Users\lap4all\Downloads\NTB - Báo Cáo Tuần (HRBP - ARD) - Tuần 32 - Final.docx'
out_path = r'C:\Users\lap4all\Downloads\NTB - Báo Cáo Tuần (HRBP - ARD) - Tuần 32 - FullUpdated.docx'
target_path = r'C:\Users\lap4all\Downloads\NTB - Báo Cáo Tuần (HRBP - ARD) - Tuần 32 - Final.docx'
orig_path = r'C:\Users\lap4all\Downloads\NTB - Báo Cáo Tuần (HRBP - ARD) - Tuần 32.docx'

path3 = r'C:\Users\lap4all\Documents\Auto report\scratch\new_image3_w31_vs_w32.png'
path21 = r'C:\Users\lap4all\Documents\Auto report\scratch\new_image21_trend_w29_w32.png'

with open(path3, 'rb') as f:
    bytes3 = f.read()
with open(path21, 'rb') as f:
    bytes21 = f.read()

doc = docx.Document(docx_path)

replaced = []
for rel in doc.part.rels.values():
    if 'image' in rel.target_ref:
        filename = os.path.basename(rel.target_ref)
        if filename == 'image3.png':
            rel.target_part._blob = bytes3
            replaced.append(filename)
        elif filename == 'image21.png':
            rel.target_part._blob = bytes21
            replaced.append(filename)

doc.save(out_path)
print(f"Saved updated file to: {out_path}")
print(f"Replaced images: {replaced}")

for dst in [target_path, orig_path]:
    try:
        doc.save(dst)
        print(f"Successfully overwritten: {dst}")
    except Exception as e:
        print(f"Could not overwrite {dst} directly (likely open in Word): {e}")
