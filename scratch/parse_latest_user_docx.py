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
        raw_title = txt
        title_name = raw_title.split('(')[0].strip()
        if '. ' in title_name: title_name = title_name.split('. ', 1)[1].strip()
        prov_name = 'TỈNH LÂM ĐỒNG'
        if '(Tỉnh Khánh Hòa)' in raw_title: prov_name = 'TỈNH KHÁNH HÒA'
        elif '(Tỉnh Ninh Thuận)' in raw_title: prov_name = 'TỈNH NINH THUẬN'
        elif '(Tỉnh Bình Thuận)' in raw_title: prov_name = 'TỈNH BÌNH THUẬN'
        elif '(Tỉnh Đắk Nông)' in raw_title: prov_name = 'TỈNH ĐẮK NÔNG'
        
        current = {'raw_heading': raw_title, 'title': title_name, 'province': prov_name, 'communes': [], 'proposal': '', 'reason': ''}
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
                raw_ward = txt.split('(Giao')[0].split('(Giao :')[0].replace('•', '').replace('-', '').strip()
                if raw_ward and raw_ward not in current['communes']:
                    current['communes'].append(raw_ward)

if current: parsed_36.append(current)

print(f"Total parsed items from latest docx: {len(parsed_36)}\n")

for item in parsed_36:
    c_str = ' + '.join(item['communes']) if item['communes'] else item['title']
    print(f"[{item['raw_heading']}]")
    print(f"   Xã cũ: {c_str}")
    print(f"   Phương án: {item['proposal']}")
    print(f"   Lý do: {item['reason']}")
    print('-'*70)
