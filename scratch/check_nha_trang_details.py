import docx, sys

sys.stdout.reconfigure(encoding='utf-8')

doc = docx.Document(r'C:\Users\lap4all\Downloads\Quy_Hoach MANG LUOI NTB.docx')

print("=== SEARCHING NHA TRANG / MO MOI / DI DOI IN DOCX ===\n")

for i, p in enumerate(doc.paragraphs):
    txt = p.text.strip()
    if not txt: continue
    
    if any(k in txt.lower() for k in ['nha trang', 'mở mới', 'di dời', 'bắc cam ranh', 'nam nha trang', 'kho mới']):
        print(f"P{i+1}: {txt}\n")
