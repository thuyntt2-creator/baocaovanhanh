import docx, sys

sys.stdout.reconfigure(encoding='utf-8')

doc = docx.Document(r'C:\Users\lap4all\Downloads\Quy_Hoach MANG LUOI NTB.docx')

print("=== ALL TEXT ABOUT DI LINH IN DOCX ===\n")

for i, p in enumerate(doc.paragraphs):
    txt = p.text.strip()
    if not txt: continue
    
    if 'di linh' in txt.lower() or 'đinh trang thượng' in txt.lower() or 'hòa ninh' in txt.lower():
        print(f"P{i+1}: {txt}\n")
