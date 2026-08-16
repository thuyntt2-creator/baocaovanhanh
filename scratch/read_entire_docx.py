import zipfile
import xml.etree.ElementTree as ET
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

docx_path = r"C:\Users\lap4all\Downloads\AOP_BCCK_Plan_new.docx"

def get_docx_all_elements(path):
    if not os.path.exists(path):
        return f"File not found: {path}"
    try:
        with zipfile.ZipFile(path) as z:
            xml_content = z.read('word/document.xml')
            root = ET.fromstring(xml_content)
            namespaces = {
                'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
            }
            
            body = root.find('w:body', namespaces)
            if body is None:
                return "No body found"
                
            out = []
            for child in body:
                tag = child.tag.split('}')[-1]
                if tag == 'p':
                    # Paragraph
                    p_text = "".join(t.text for t in child.findall('.//w:t', namespaces) if t.text)
                    out.append(f"[P] {p_text}")
                elif tag == 'tbl':
                    # Table
                    out.append("[Table]")
                    rows = child.findall('w:tr', namespaces)
                    for r_idx, r in enumerate(rows):
                        cells = r.findall('w:tc', namespaces)
                        cell_texts = []
                        for cell in cells:
                            cell_text = "".join(t.text for t in cell.findall('.//w:t', namespaces) if t.text)
                            cell_texts.append(cell_text.strip())
                        out.append(f"  Row {r_idx}: " + " | ".join(cell_texts))
            return "\n".join(out)
    except Exception as e:
        return f"Error: {e}"

print(get_docx_all_elements(docx_path))

