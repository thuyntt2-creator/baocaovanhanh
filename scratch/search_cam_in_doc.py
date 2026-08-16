import zipfile
import xml.etree.ElementTree as ET
import sys

sys.stdout.reconfigure(encoding='utf-8')
target_path = r"C:\Users\lap4all\Downloads\Bao_Cao_Quy_Hoach_ MANG LUOI NTB.docx"

with zipfile.ZipFile(target_path) as z:
    root = ET.fromstring(z.read('word/document.xml'))
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    body = root.find('w:body', ns)
    for child in body:
        tag = child.tag.split('}')[-1]
        if tag == 'p':
            t = ''.join(e.text for e in child.findall('.//w:t', ns) if e.text).strip()
            if 'Cam' in t or 'Khánh' in t:
                print('P:', t[:120])
