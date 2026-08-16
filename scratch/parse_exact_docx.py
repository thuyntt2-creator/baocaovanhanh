import docx, sys

sys.stdout.reconfigure(encoding='utf-8')

doc = docx.Document(r'C:\Users\lap4all\Downloads\Quy_Hoach MANG LUOI NTB.docx')

parsed_36 = []
current = None

for p in doc.paragraphs:
    txt = p.text.strip()
    if not txt: continue
    
    is_heading = False
    for n in range(1, 37):
        if txt.startswith(f'{n}. '):
            is_heading = True
            break
            
    if is_heading:
        if current: parsed_36.append(current)
        current = {'heading': txt, 'commune_bullets': [], 'proposal': '', 'reason': ''}
    elif current:
        if 'ĐỀ XUẤT PHƯƠNG ÁN' in txt:
            current['proposal'] = txt.replace('❖ ĐỀ XUẤT PHƯƠNG ÁN CỦA AM:', '').replace('❖ ĐỀ XUẤT PHƯƠNG ÁN:', '').strip()
        elif 'LÝ DO VÀ GIẢI THÍCH' in txt:
            current['reason'] = txt.replace('❖ LÝ DO VÀ GIẢI THÍCH NGUYÊN NHÂN CHI TIẾT TỪ AM:', '').replace('❖ LÝ DO VÀ GIẢI THÍCH NGUYÊN NHÂN CHI TIẾT:', '').strip()
        elif current['proposal'] and not current['reason'] and not txt.startswith('❖'):
            current['proposal'] += ' ' + txt
        elif current['reason'] and not txt.startswith('❖'):
            current['reason'] += ' ' + txt
        elif not current['proposal'] and not txt.startswith('❖'):
            if ('Giao:' in txt or 'Lấy:' in txt or 'BC:' in txt or 'Phường' in txt or 'Xã' in txt or 'Thị trấn' in txt) and not txt.startswith('TỔNG SẢN LƯỢNG') and not txt.startswith('Các BC hiện') and not txt.startswith('Mã Xã') and not txt.startswith('Tỷ lệ'):
                # Extract clean commune/ward name
                raw_name = txt.split('(Giao')[0].split('(Giao :')[0].replace('•', '').replace('-', '').strip()
                if raw_name and raw_name not in current['commune_bullets']:
                    current['commune_bullets'].append(raw_name)

if current: parsed_36.append(current)

print(f"Total parsed items: {len(parsed_36)}\n")

for idx, item in enumerate(parsed_36, 1):
    communes_str = ' + '.join(item['commune_bullets'])
    print(f"{idx}. {item['heading']}")
    print(f"   Xã cũ: {communes_str}")
    print(f"   Phương án: {item['proposal'][:100]}")
    print(f"   Lý do: {item['reason'][:100]}")
    print('-'*60)
