import zipfile
import xml.etree.ElementTree as ET
import sys

sys.stdout.reconfigure(encoding='utf-8')

target_path = r"c:\Users\lap4all\Documents\Auto report\Bao_Cao_Quy_Hoach_Buu_Cuc_NTB_Theo_DVHC_Moi_Co_Hinh_AM.docx"

with zipfile.ZipFile(target_path) as z:
    root = ET.fromstring(z.read('word/document.xml'))
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
          'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
          'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
    body = root.find('w:body', ns)
    
    last_text = ""
    for child in body:
        tag = child.tag.split('}')[-1]
        if tag == 'p':
            t = ''.join(e.text for e in child.findall('.//w:t', ns) if e.text).strip()
            if t:
                last_text = t
            # check drawing / blip
            blips = child.findall('.//a:blip', ns)
            for b in blips:
                r_id = b.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                print(f"IMAGE (rId: {r_id}) after text: {last_text[:120]}")
