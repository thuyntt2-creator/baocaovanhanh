# -*- coding: utf-8 -*-
import sys, os, docx

sys.stdout.reconfigure(encoding='utf-8')

guideline_doc_path = r'C:\Users\lap4all\Downloads\HƯỚNG DẪN XÂY DỰNG KẾ HOẠCH VẬN HÀNH CHO CÁC VÙNG - EVENT.docx'

doc_g = docx.Document(guideline_doc_path)

print('=== FULL TEXT OF COMPANY GUIDELINE ===')
for idx, p in enumerate(doc_g.paragraphs):
    txt = p.text.strip()
    if txt:
        style = p.style.name if p.style else 'Normal'
        print(f'[{style}] {txt}')

print('\n=== ALL TABLES IN COMPANY GUIDELINE ===')
for t_idx, table in enumerate(doc_g.tables):
    rows_cnt = len(table.rows)
    cols_cnt = len(table.columns)
    print(f'\n--- Table {t_idx+1} ({rows_cnt} rows x {cols_cnt} cols) ---')
    for r_i in range(rows_cnt):
        row_txt = [cell.text.replace('\n', ' ').strip() for cell in table.rows[r_i].cells]
        print(f'Row {r_i}:', ' | '.join(row_txt))
