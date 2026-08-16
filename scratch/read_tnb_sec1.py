# -*- coding: utf-8 -*-
import sys, os, docx

sys.stdout.reconfigure(encoding='utf-8')

tnb_sample_path = r'C:\Users\lap4all\Downloads\Copy of Copy of TNB Kế hoạch Event 7.7.docx'
doc = docx.Document(tnb_sample_path)

print('=== TNB SECTION I. MỤC TIÊU ===')
in_sec1 = False
for p in doc.paragraphs:
    txt = p.text.strip()
    if 'I. MỤC TIÊU' in txt.upper() or 'MỤC TIÊU' in txt.upper():
        in_sec1 = True
    elif 'II. PHÂN TÍCH' in txt.upper():
        in_sec1 = False
        
    if in_sec1 and txt:
        print(f'[{p.style.name}] {txt}')
