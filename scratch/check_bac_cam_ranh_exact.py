import docx, sys

sys.stdout.reconfigure(encoding='utf-8')

doc = docx.Document(r'C:\Users\lap4all\Downloads\Quy_Hoach MANG LUOI NTB.docx')

print("=== CHECKING BẮC CAM RANH & CAM NGHĨA IN DOCX ===\n")

for i, p in enumerate(doc.paragraphs):
    txt = p.text.strip()
    if not txt: continue
    
    if 'bắc cam ranh' in txt.lower() or 'cam nghĩa' in txt.lower() or 'bac cam ranh' in txt.lower():
        print(f"P{i+1}: {txt}\n")
