import docx, os, sys

sys.stdout.reconfigure(encoding='utf-8')

src_docx = r'C:\Users\lap4all\Downloads\Quy_Hoach MANG LUOI NTB.docx'
doc = docx.Document(src_docx)

# Find all "Hình X:" text in paragraphs and their exact line numbers
hinh_captions = []
for i, p in enumerate(doc.paragraphs):
    txt = p.text.strip()
    if 'Hình' in txt:
        hinh_captions.append((i+1, txt))

print("=== ALL HÌNH CAPTIONS IN DOCX ===")
for p_idx, cap in hinh_captions:
    print(f"P{p_idx}: {cap}")
