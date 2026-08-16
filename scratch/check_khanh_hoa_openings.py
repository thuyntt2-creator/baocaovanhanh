import docx, sys

sys.stdout.reconfigure(encoding='utf-8')

doc = docx.Document(r'C:\Users\lap4all\Downloads\Quy_Hoach MANG LUOI NTB.docx')

print("=== ALL KHÁNH HÒA POST OFFICE OPENINGS / SPLITS / RELOCATIONS ===\n")

for i, p in enumerate(doc.paragraphs):
    txt = p.text.strip()
    if not txt: continue
    
    if any(k in txt.lower() for k in ['khánh hòa', 'nha trang', 'cam ranh', 'nam nha trang', 'tây nha trang', 'mở mới', 'tách mới']):
        if any(w in txt.lower() for w in ['mở mới', 'tách mới', 'mở bưu cục', 'bưu cục mới', 'kho mới']):
            print(f"P{i+1}: {txt}\n")
