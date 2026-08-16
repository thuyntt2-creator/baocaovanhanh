import docx
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

docx_path = r'C:\Users\lap4all\Downloads\NTB - Báo Cáo Tuần (HRBP - ARD) - Tuần 32 - Final.docx'
out_path = r'C:\Users\lap4all\Downloads\NTB - Báo Cáo Tuần (HRBP - ARD) - Tuần 32 - FullUpdated.docx'
target_path = r'C:\Users\lap4all\Downloads\NTB - Báo Cáo Tuần (HRBP - ARD) - Tuần 32 - Final.docx'

new_img3_path = r'C:\Users\lap4all\Documents\Auto report\scratch\new_image3_w31_vs_w32.png'

with open(new_img3_path, 'rb') as f:
    new_img3_bytes = f.read()

doc = docx.Document(docx_path)

replaced_count = 0
for rel in doc.part.rels.values():
    if 'image' in rel.target_ref:
        filename = os.path.basename(rel.target_ref)
        if filename == 'image3.png':
            rel.target_part._blob = new_img3_bytes
            replaced_count += 1
            print(f"Successfully replaced blob for {filename}!")

doc.save(out_path)
print(f"Saved docx with replaced image3.png to {out_path}. Replaced count: {replaced_count}")

try:
    doc.save(target_path)
    print(f"Successfully overwritten {target_path}!")
except Exception as e:
    print(f"Could not overwrite {target_path} directly (open in Word): {e}")
