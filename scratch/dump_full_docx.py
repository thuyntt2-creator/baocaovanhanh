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
    
    idx = 0
    for child in body:
        tag = child.tag.split('}')[-1]
        if tag == 'p':
            t = ''.join(e.text for e in child.findall('.//w:t', ns) if e.text)
            if t.strip():
                print(f"P[{idx}]: {t}")
        elif tag == 'tbl':
            print(f"--- TABLE[{idx}] START ---")
            for row in child.findall('.//w:tr', ns):
                cells = []
                for cell in row.findall('.//w:tc', ns):
                    ctext = ''.join(e.text for e in cell.findall('.//w:t', ns) if e.text)
                    cells.append(ctext.strip().replace('\n', ' '))
                print(' | '.join(cells))
            print(f"--- TABLE[{idx}] END ---")
        idx += 1
