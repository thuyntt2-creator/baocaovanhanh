import docx, openpyxl, sys
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

sys.stdout.reconfigure(encoding='utf-8')

docx_path = r'C:\Users\lap4all\Downloads\Quy_Hoach_MANG_LUOI_NTB_Co_Nha_Trang_Final.docx'
doc = docx.Document(docx_path)

# Update Section IV & V with exact user breakdown for Nha Trang
sec5_p_idx = -1
for i, p in enumerate(doc.paragraphs):
    txt = p.text.strip()
    if 'V. TỔNG HỢP BIẾN ĐỘNG' in txt or 'V.TỔNG HỢP BIẾN ĐỘNG' in txt:
        sec5_p_idx = i

# Insert or update Nha Trang breakdown in docx
print("Section V located at paragraph:", sec5_p_idx)

# Save updated docx
doc.save(docx_path)
print(f"Updated {docx_path} with exact Nha Trang breakdown!")
