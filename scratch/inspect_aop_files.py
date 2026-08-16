import sys
import zipfile
import xml.etree.ElementTree as ET
import openpyxl
import os

# Set stdout encoding to UTF-8
sys.stdout.reconfigure(encoding='utf-8')

docx_path = r"C:\Users\lap4all\Downloads\AOP_BCCK_Plan_new.docx"
xlsx_aop_path = r"C:\Users\lap4all\Downloads\AOP_Hang_NTB_T7-T12_2026.xlsx"
xlsx_config_path = r"C:\Users\lap4all\Downloads\config_psbba_NTB.xlsx"
xlsx_compare_path = r"C:\Users\lap4all\Downloads\[NTB] So sánh topline H2 Mới - Cũ.xlsx"

def get_docx_text(path):
    if not os.path.exists(path):
        return f"File not found: {path}"
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
        return f"Error reading {path}: {str(e)}"

def inspect_xlsx(path):
    if not os.path.exists(path):
        return f"File not found: {path}"
    try:
        wb = openpyxl.load_workbook(path, read_only=True)
        return f"Sheets in {os.path.basename(path)}: {wb.sheetnames}"
    except Exception as e:
        return f"Error reading {path}: {str(e)}"

print("=== DOCX TEXT ===")
print(get_docx_text(docx_path)[:3000]) # print first 3000 chars

print("\n=== XLSX INSPECTION ===")
print(inspect_xlsx(xlsx_aop_path))
print(inspect_xlsx(xlsx_config_path))
print(inspect_xlsx(xlsx_compare_path))
