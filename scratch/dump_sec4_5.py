import zipfile
import xml.etree.ElementTree as ET
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

target_path = r"C:\Users\lap4all\Downloads\Bao_Cao_Quy_Hoach_ MANG LUOI NTB.docx"

with zipfile.ZipFile(target_path) as z:
    xml_content = z.read('word/document.xml')
    root = ET.fromstring(xml_content)
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    body = root.find('w:body', ns)
    
    recording = False
    for child in body:
        tag = child.tag.split('}')[-1]
        if tag == 'p':
            t = ''.join(e.text for e in child.findall('.//w:t', ns) if e.text)
            if 'IV.' in t or 'PHẦN 4' in t or 'YÊU CẦU 3' in t:
                recording = True
            if recording:
                print('P:', t)
        elif tag == 'tbl' and recording:
            print('--- TABLE START ---')
            for row in child.findall('.//w:tr', ns):
                cells = []
                for cell in row.findall('.//w:tc', ns):
                    ctext = ''.join(e.text for e in cell.findall('.//w:t', ns) if e.text)
                    cells.append(ctext.strip().replace('\n', ' '))
                print(' | '.join(cells))
            print('--- TABLE END ---')
