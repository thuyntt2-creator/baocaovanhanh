import zipfile
import xml.etree.ElementTree as ET
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

docx_files = [
    r"C:\Users\lap4all\Downloads\AOP_BCCK_Plan_new.docx",
    r"C:\Users\lap4all\Downloads\AOP_Hang_Nang_NTB_T7-T12_2026_FINAL.docx",
    r"C:\Users\lap4all\Downloads\AOP_BCCK_Plan_NTB_new.docx"
]

def get_docx_text(path):
    if not os.path.exists(path):
        return ""
    try:
        with zipfile.ZipFile(path) as z:
            xml_content = z.read('word/document.xml')
            root = ET.fromstring(xml_content)
            namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
            paragraphs = []
            for p in root.findall('.//w:p', namespaces):
                p_text = "".join(t.text for t in p.findall('.//w:t', namespaces) if t.text)
                if p_text:
                    paragraphs.append(p_text)
            return "\n".join(paragraphs)
    except Exception as e:
        return f"Error: {e}"

for fpath in docx_files:
    if os.path.exists(fpath):
        print(f"\n==========================================")
        print(f"File: {os.path.basename(fpath)}")
        print(f"==========================================")
        text = get_docx_text(fpath)
        
        # Search for lines containing keywords
        lines = text.split("\n")
        print(f"Total lines: {len(lines)}")
        
        # Search for Di Linh or hỗn hợp or chuyên biệt
        for idx, line in enumerate(lines):
            if "hỗn hợp" in line or "Di Linh" in line or "chuyên biệt" in line or "chuyên trách" in line or "độc lập" in line:
                print(f"  Line {idx+1}: {line}")
