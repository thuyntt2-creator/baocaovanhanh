import docx, sys

sys.stdout.reconfigure(encoding='utf-8')

doc = docx.Document(r'C:\Users\lap4all\Downloads\Quy_Hoach MANG LUOI NTB.docx')

print("=== CHECKING BẢO LỘC 1 IN DOCX ===\n")

for i, p in enumerate(doc.paragraphs):
    txt = p.text.strip()
    if not txt: continue
    
    if 'bảo lộc' in txt.lower() or 'blao' in txt.lower() or '1 bảo lộc' in txt.lower() or 'b\'lao' in txt.lower():
        print(f"P{i+1}: {txt}\n")
