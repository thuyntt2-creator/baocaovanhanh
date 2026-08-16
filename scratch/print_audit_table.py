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
        current = {'heading': txt, 'proposal': '', 'reason': ''}
    elif current:
        if 'ĐỀ XUẤT PHƯƠNG ÁN' in txt:
            current['proposal'] = txt.replace('❖ ĐỀ XUẤT PHƯƠNG ÁN CỦA AM:', '').replace('❖ ĐỀ XUẤT PHƯƠNG ÁN:', '').strip()
        elif 'LÝ DO VÀ GIẢI THÍCH' in txt:
            current['reason'] = txt.replace('❖ LÝ DO VÀ GIẢI THÍCH NGUYÊN NHÂN CHI TIẾT TỪ AM:', '').replace('❖ LÝ DO VÀ GIẢI THÍCH NGUYÊN NHÂN CHI TIẾT:', '').strip()
        elif current['proposal'] and not current['reason'] and not txt.startswith('❖'):
            current['proposal'] += ' ' + txt
        elif current['reason'] and not txt.startswith('❖'):
            current['reason'] += ' ' + txt

if current: parsed_36.append(current)

print("=== COMPLETE SYSTEMATIC AUDIT TABLE FOR MỞ MỚI / TÁCH MỚI / DI DỜI / ĐÓNG CỬA ===\n")

open_split_items = []
move_expand_items = []
close_items = []

for item in parsed_36:
    text = (item['proposal'] + ' ' + item['reason']).upper()
    h = item['heading']
    
    # Check open / split
    if any(k in text for k in ['TÁCH MỚI', 'MỞ MỚI', 'TÁCH BƯU CỤC', 'MỞ BƯU CỤC']):
        open_split_items.append((h, item['proposal']))
        
    # Check move / expand
    if any(k in text for k in ['DI DỜI', 'MỞ RỘNG MẶT BẰNG', 'TIM MB MÓI', 'MẶT BẰNG MÓI', 'TÌM MB MỚI', 'MB MỚI']):
        move_expand_items.append((h, item['proposal']))
        
    # Check close
    if any(k in text for k in ['ĐÓNG CỬA', 'BỎ BƯU CỤC', 'BỎ BC']):
        close_items.append((h, item['proposal']))

print("📌 1. DANH SÁCH BƯU CỤC MỞ MỚI / TÁCH MỚI:")
for h, prop in open_split_items:
    print(f"  • [{h}]: {prop}\n")

print("\n📌 2. DANH SÁCH BƯU CỤC DI DỜI / MỞ RỘNG MẶT BẰNG KHO:")
for h, prop in move_expand_items:
    print(f"  • [{h}]: {prop}\n")

print("\n📌 3. DANH SÁCH BƯU CỤC ĐÓNG CỬA / GỘP KHO:")
for h, prop in close_items:
    print(f"  • [{h}]: {prop}\n")

