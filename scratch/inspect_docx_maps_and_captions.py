import docx, os, sys

sys.stdout.reconfigure(encoding='utf-8')

src_docx = r'C:\Users\lap4all\Downloads\Quy_Hoach MANG LUOI NTB.docx'
doc = docx.Document(src_docx)

print("=== INSPECTING ALL MAP PARAGRAPHS AND HINH CAPTIONS IN DOCX ===\n")

for i, p in enumerate(doc.paragraphs):
    txt = p.text.strip()
    if not txt: continue
    
    # Check if paragraph has images or "Hình" text
    has_img = False
    for run in p.runs:
        if run._r.xpath('.//w:drawing'):
            has_img = True
            break
            
    if has_img or 'Hình' in txt:
        print(f"P{i+1}: Text='{txt}' | HasImage={has_img}")

