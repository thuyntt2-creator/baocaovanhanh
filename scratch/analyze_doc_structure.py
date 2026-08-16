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
    
    current_heading = "HEADER"
    lines_by_section = {}
    
    for child in body:
        tag = child.tag.split('}')[-1]
        if tag == 'p':
            t = ''.join(e.text for e in child.findall('.//w:t', ns) if e.text).strip()
            if not t:
                continue
            if t.startswith('I.') or t.startswith('II.') or t.startswith('III.') or t.startswith('IV.') or t.startswith('V.'):
                current_heading = t
                lines_by_section[current_heading] = []
            else:
                if current_heading not in lines_by_section:
                    lines_by_section[current_heading] = []
                lines_by_section[current_heading].append(('p', t))
        elif tag == 'tbl':
            if current_heading not in lines_by_section:
                lines_by_section[current_heading] = []
            lines_by_section[current_heading].append(('tbl', 'TABLE'))

for section, contents in lines_by_section.items():
    print(f"=== {section} ===")
    print(f"Total elements: {len(contents)}")
    for kind, text in contents[:10]: # print first few
        print(f"  [{kind}] {text[:100]}")
    print()
