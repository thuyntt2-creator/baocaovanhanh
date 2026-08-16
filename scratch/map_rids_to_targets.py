import zipfile
import xml.etree.ElementTree as ET
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

target_path = r"c:\Users\lap4all\Documents\Auto report\Bao_Cao_Quy_Hoach_Buu_Cuc_NTB_Theo_DVHC_Moi_Co_Hinh_AM.docx"

with zipfile.ZipFile(target_path) as z:
    rels_xml = z.read('word/_rels/document.xml.rels')
    root = ET.fromstring(rels_xml)
    ns = {'r': 'http://schemas.openxmlformats.org/package/2006/relationships'}
    
    r_map = {}
    for rel in root.findall('.//r:Relationship', ns):
        rid = rel.attrib.get('Id')
        target = rel.attrib.get('Target')
        r_map[rid] = target
        
    for rid in ['rId9', 'rId10', 'rId11', 'rId12', 'rId13', 'rId14', 'rId15', 'rId16', 'rId17']:
        print(f"{rid} -> {r_map.get(rid)}")
