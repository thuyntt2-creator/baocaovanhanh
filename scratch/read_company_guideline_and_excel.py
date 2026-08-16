# -*- coding: utf-8 -*-
import sys, os, docx, openpyxl, pandas as pd

sys.stdout.reconfigure(encoding='utf-8')

guideline_doc_path = r'C:\Users\lap4all\Downloads\HƯỚNG DẪN XÂY DỰNG KẾ HOẠCH VẬN HÀNH CHO CÁC VÙNG - EVENT.docx'
ntb_excel_path = r'C:\Users\lap4all\Downloads\config_psbba_NTB.xlsx'

print('=== 1. PARSING COMPANY GUIDELINE DOCX ===')
if os.path.exists(guideline_doc_path):
    doc_g = docx.Document(guideline_doc_path)
    print(f'Total Paragraphs: {len(doc_g.paragraphs)}')
    print(f'Total Tables: {len(doc_g.tables)}')
    
    print('\n--- GUIDELINE PARAGRAPHS & HEADINGS ---')
    for idx, p in enumerate(doc_g.paragraphs):
        txt = p.text.strip()
        if txt:
            style = p.style.name if p.style else 'Normal'
            print(f'P{idx:03d} [{style}]: {txt}')
            
    print('\n--- GUIDELINE TABLES ---')
    for t_idx, table in enumerate(doc_g.tables):
        rows_cnt = len(table.rows)
        cols_cnt = len(table.columns)
        print(f'\nTable {t_idx+1} ({rows_cnt} rows x {cols_cnt} cols):')
        for r_i in range(min(5, rows_cnt)):
            row_txt = [cell.text.replace('\n', ' ').strip() for cell in table.rows[r_i].cells]
            print(f'  Row {r_i}:', ' | '.join(row_txt))
else:
    print(f'Guideline file not found: {guideline_doc_path}')

print('\n=== 2. PARSING UPDATED CONFIG_PSBBA_NTB.XLSX SHEETS ===')
if os.path.exists(ntb_excel_path):
    xl = pd.ExcelFile(ntb_excel_path)
    print('Sheet names in config_psbba_NTB.xlsx:', xl.sheet_names)
    
    for s_name in xl.sheet_names:
        if any(k in s_name.lower() for k in ['nhân sự', 'nhan_su', 'nhansu', 'bất ổn', 'bat_on', 'baton', 'thủy']):
            print(f'\n--- Sheet: {s_name} ---')
            df_s = pd.read_excel(xl, sheet_name=s_name)
            print('Shape:', df_s.shape)
            print('Columns:', df_s.columns.tolist())
            print('Head 5 rows:')
            print(df_s.head(5))
