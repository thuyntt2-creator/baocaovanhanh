import os
import glob
import sys
import docx

sys.stdout.reconfigure(encoding='utf-8')

downloads_dir = r"C:\Users\lap4all\Downloads"

# Check docx guide
docx_path = os.path.join(downloads_dir, "Huong_dan_AOP_Hang_Nang_cho_GDV.docx")
if os.path.exists(docx_path):
    print("=== Reading Huong_dan_AOP_Hang_Nang_cho_GDV.docx ===")
    doc = docx.Document(docx_path)
    for p in doc.paragraphs:
        txt = p.text
        if any(term in txt.lower() for term in ["setup", "mở mới", "di dời", "khoán"]):
            print(f"  {txt}")
            
# Check html guide
html_path = os.path.join(downloads_dir, "Huong_dan_chi_tiet_lap_AOP_Hang_Nang.html")
if os.path.exists(html_path):
    print("\n=== Reading Huong_dan_chi_tiet_lap_AOP_Hang_Nang.html ===")
    with open(html_path, 'r', encoding='utf-8', errors='ignore') as f:
        html = f.read()
        lines = html.split('\n')
        for idx, line in enumerate(lines):
            if any(term in line.lower() for term in ["setup", "mở mới", "di dời"]):
                print(f"  Line {idx}: {line.strip()[:150]}")
