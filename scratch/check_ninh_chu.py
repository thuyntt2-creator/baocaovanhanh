import docx, sys

sys.stdout.reconfigure(encoding='utf-8')

doc = docx.Document(r'C:\Users\lap4all\Downloads\Quy_Hoach MANG LUOI NTB.docx')

print("=== CHECKING NINH CHỬ IN DOCX ===\n")

for i, p in enumerate(doc.paragraphs):
    txt = p.text.strip()
    if not txt: continue
    
    if 'ninh chử' in txt.lower() or 'ninh chu' in txt.lower():
        print(f"P{i+1}: {txt}\n")
