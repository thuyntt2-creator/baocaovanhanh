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
        current = {'num': txt.split('.')[0].strip(), 'title': txt, 'proposal': '', 'reason': ''}
    elif current:
        if 'ĐỀ XUẤT PHƯƠNG ÁN' in txt:
            current['proposal'] = txt
        elif 'LÝ DO VÀ GIẢI THÍCH' in txt:
            current['reason'] = txt
        elif current['proposal'] and not current['reason'] and not txt.startswith('❖'):
            current['proposal'] += ' ' + txt

if current: parsed_36.append(current)

# Categorize items
mo_moi = []
di_doi = []
dong_cua = []
gop_tuyen = []
giu_nguyen = []

for item in parsed_36:
    text = (item['proposal'] + ' ' + item['reason']).upper()
    t = item['title']
    
    if 'TÁCH MỚI' in text or 'MỞ MỚI' in text or 'TÁCH BƯU CỤC' in text or 'MỞ BƯU CỤC' in text:
        mo_moi.append(t)
    elif 'DI DỜI' in text or 'MỞ RỘNG MẶT BẰNG' in text:
        di_doi.append(t)
    elif 'ĐÓNG CỬA' in text:
        dong_cua.append(t)
    elif 'GIỮ NGUYÊN' in text or 'TẠM THỜI GIỮ NGUYÊN' in text:
        giu_nguyen.append(t)
    else:
        gop_tuyen.append(t)

print("=== VERIFICATION RESULTS FROM ALL 36 ITEMS ===")
print(f"\n1. Mở mới / Tách BC ({len(mo_moi)} mục):")
for x in mo_moi: print("   -", x)

print(f"\n2. Di dời / Mở rộng ({len(di_doi)} mục):")
for x in di_doi: print("   -", x)

print(f"\n3. Đóng cửa BC ({len(dong_cua)} mục):")
for x in dong_cua: print("   -", x)

print(f"\n4. Giữ nguyên vận hành ({len(giu_nguyen)} mục):")
for x in giu_nguyen: print("   -", x)

print(f"\n5. Gộp tuyến / Tối ưu phân vùng ({len(gop_tuyen)} mục):")
for x in gop_tuyen: print("   -", x)
