import os
import glob
import sys
import docx

sys.stdout.reconfigure(encoding='utf-8')

downloads_dir = r"C:\Users\lap4all\Downloads"

docx_path = os.path.join(downloads_dir, "Huong_dan_AOP_Hang_Nang_cho_GDV.docx")
if os.path.exists(docx_path):
    print("=== paragraphs in docx ===")
    doc = docx.Document(docx_path)
    for idx, p in enumerate(doc.paragraphs):
        txt = p.text
        if any(term in txt.lower() for term in ["setup", "mở mới", "mở", "di dời", "dời", "đức linh"]):
            print(f"P{idx}: {txt}")

html_path = os.path.join(downloads_dir, "Huong_dan_chi_tiet_lap_AOP_Hang_Nang.html")
if os.path.exists(html_path):
    print("\n=== lines in html ===")
    with open(html_path, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()
        lines = content.split('\n')
        for idx, line in enumerate(lines):
            if any(term in line.lower() for term in ["setup", "mở mới", "di dời", "dời", "đức linh"]):
                print(f"L{idx}: {line.strip()}")
