# -*- coding: utf-8 -*-
import sys, os, docx

sys.stdout.reconfigure(encoding='utf-8')

tnb_sample_path = r'C:\Users\lap4all\Downloads\Copy of Copy of TNB Kế hoạch Event 7.7.docx'

if not os.path.exists(tnb_sample_path):
    print(f'File not found: {tnb_sample_path}')
    sys.exit(1)

doc = docx.Document(tnb_sample_path)

print('=== TNB SAMPLE DOCUMENT STRUCTURE & CONTENT ===')
print(f'Total Paragraphs: {len(doc.paragraphs)}')
print(f'Total Tables: {len(doc.tables)}')
print(f'Total Inline Shapes (Images/Charts): {len(doc.inline_shapes)}')

print('\n--- ALL PARAGRAPHS & HEADINGS ---')
for idx, p in enumerate(doc.paragraphs):
    txt = p.text.strip()
    if txt:
        style = p.style.name if p.style else 'Normal'
        print(f'P{idx:03d} [{style}]: {txt}')

print('\n--- ALL TABLES SUMMARY ---')
for t_idx, table in enumerate(doc.tables):
    rows_cnt = len(table.rows)
    cols_cnt = len(table.columns)
    first_row = [cell.text.replace('\n', ' ').strip() for cell in table.rows[0].cells]
    print(f'\nTable {t_idx+1} ({rows_cnt} rows x {cols_cnt} cols):')
    print('Header:', ' | '.join(first_row[:6]))
    if rows_cnt > 1:
        row1 = [cell.text.replace('\n', ' ').strip() for cell in table.rows[1].cells]
        print('Row 1 :', ' | '.join(row1[:6]))
