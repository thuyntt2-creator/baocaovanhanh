import zipfile
import xml.etree.ElementTree as ET
import os
import glob
import sys

sys.stdout.reconfigure(encoding='utf-8')

download_dir = r"C:\Users\lap4all\Downloads"
print("Scanning Downloads folder:")
matches = []
for f in os.listdir(download_dir):
    if f.lower().endswith(".docx"):
        print("DOCX file:", f)
        if "ntb" in f.lower() or "quy_hoach" in f.lower() or "mang" in f.lower() or "bao_cao" in f.lower():
            matches.append(os.path.join(download_dir, f))

print("\nMatching files:", matches)

target_path = r"C:\Users\lap4all\Downloads\Bao_Cao_Quy_Hoach_ MANG LUOI NTB.docx"
if not os.path.exists(target_path):
    for m in matches:
        if "mang luoi" in m.lower() or "mang_luoi" in m.lower() or "mang" in m.lower():
            target_path = m
            break

print(f"\n--- Reading {target_path} ---")
if os.path.exists(target_path):
    with zipfile.ZipFile(target_path) as z:
        xml_content = z.read('word/document.xml')
        root = ET.fromstring(xml_content)
        ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        body = root.find('w:body', ns)
        for child in body:
            tag = child.tag.split('}')[-1]
            if tag == 'p':
                t = ''.join(e.text for e in child.findall('.//w:t', ns) if e.text)
                if t.strip():
                    print('P:', t)
            elif tag == 'tbl':
                print('--- TABLE START ---')
                for row in child.findall('.//w:tr', ns):
                    cells = []
                    for cell in row.findall('.//w:tc', ns):
                        ctext = ''.join(e.text for e in cell.findall('.//w:t', ns) if e.text)
                        cells.append(ctext.strip().replace('\n', ' '))
                    print(' | '.join(cells))
                print('--- TABLE END ---')
