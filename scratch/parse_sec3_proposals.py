import zipfile
import xml.etree.ElementTree as ET
import sys

sys.stdout.reconfigure(encoding='utf-8')
target_path = r"C:\Users\lap4all\Downloads\Bao_Cao_Quy_Hoach_ MANG LUOI NTB.docx"

with zipfile.ZipFile(target_path) as z:
    xml_content = z.read('word/document.xml')
    root = ET.fromstring(xml_content)
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    body = root.find('w:body', ns)
    
    communes = []
    curr_c = None
    for child in body:
        if child.tag.endswith('p'):
            t = ''.join(e.text for e in child.findall('.//w:t', ns) if e.text).strip()
            if t and t[0].isdigit() and '.' in t[:4] and 'Xã' in t or 'Phường' in t:
                if curr_c:
                    communes.append(curr_c)
                curr_c = {'title': t, 'dexuat': '', 'lydo': ''}
            elif curr_c:
                if 'ĐỀ XUẤT' in t:
                    curr_c['dexuat'] = t
                elif 'LÝ DO' in t:
                    curr_c['lydo'] = t
    if curr_c:
        communes.append(curr_c)

print(f"Total communes parsed in Sec III: {len(communes)}")
for idx, c in enumerate(communes, 1):
    print(f"{idx}. {c['title']}")
    print(f"   Proposal: {c['dexuat']}")
