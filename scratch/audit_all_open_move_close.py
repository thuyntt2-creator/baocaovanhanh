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
        current = {'num': txt.split('.')[0].strip(), 'raw_heading': txt, 'lines': []}
    elif current:
        if txt.startswith(('1.', '2.', '3.', '4.', '5.', '6.', '7.', '8.', '9.', '10.', '11.', '12.', '13.', '14.', '15.', '16.', '17.', '18.', '19.', '20.', '21.', '22.', '23.', '24.', '25.', '26.', '27.', '28.', '29.', '30.', '31.', '32.', '33.', '34.', '35.', '36.', 'IV.', 'V.')):
            pass
        else:
            current['lines'].append(txt)

if current: parsed_36.append(current)

print("=== DETAILED AUDIT OF ALL 36 ITEMS FOR MO MOI / TACH MOI / DI DOI / DONG CUA ===\n")

open_split_list = []
move_expand_list = []
close_list = []

for item in parsed_36:
    full_text = ' '.join(item['lines'])
    u_text = full_text.upper()
    h = item['raw_heading']
    
    has_open = any(k in u_text for k in ['TÁCH MỚI', 'MỞ MỚI', 'TÁCH BƯU CỤC', 'MỞ BƯU CỤC', 'TÁCH BC', 'MỞ BC'])
    has_move = any(k in u_text for k in ['DI DỜI', 'MỞ RỘNG MẶT BẰNG', 'TIM MB MÓI', 'MẶT BẰNG MÓI', 'TÌM MB MỚI'])
    has_close = any(k in u_text for k in ['ĐÓNG CỬA', 'BỎ BƯU CỤC', 'BỎ BC'])
    
    print(f"[{h}]")
    if has_open:
        print(f"  👉 [MỞ MỚI / TÁCH MỚI]: {full_text[:180]}...")
        open_split_list.append((h, full_text))
    if has_move:
        print(f"  🚚 [DI DỜI / MỞ RỘNG MB]: {full_text[:180]}...")
        move_expand_list.append((h, full_text))
    if has_close:
        print(f"  ❌ [ĐÓNG CỬA / GỘP KHO]: {full_text[:180]}...")
        close_list.append((h, full_text))
    if not (has_open or has_move or has_close):
        print(f"  ℹ️ [GỘP TUYẾN / GIỮ NGUYÊN]: {full_text[:120]}...")
    print('-'*80)

print(f"\nSummary counts from Audit:")
print(f"- Total Mở mới / Tách mới: {len(open_split_list)}")
print(f"- Total Di dời / Mở rộng MB: {len(move_expand_list)}")
print(f"- Total Đóng cửa / Gộp kho: {len(close_list)}")
