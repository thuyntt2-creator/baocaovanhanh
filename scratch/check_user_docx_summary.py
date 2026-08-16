import docx, sys

sys.stdout.reconfigure(encoding='utf-8')

doc = docx.Document(r'C:\Users\lap4all\Downloads\Quy_Hoach MANG LUOI NTB.docx')

print("=== SUMMARY SECTIONS FROM USER DOCX ===\n")

in_summary = False
for i, p in enumerate(doc.paragraphs):
    txt = p.text.strip()
    if not txt: continue
    
    if 'IV.' in txt or 'V.' in txt or 'TỔNG HỢP' in txt or 'Kế hoạch Mở mới' in txt or 'ĐÁNH GIÁ' in txt or '4. Phường' in txt:
        in_summary = True
        
    if in_summary:
        print(f"P{i+1}: {txt}\n")
