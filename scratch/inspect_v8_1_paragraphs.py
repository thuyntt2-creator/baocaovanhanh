import docx
import sys

sys.stdout.reconfigure(encoding='utf-8')

docx_path = r"C:\Users\lap4all\Downloads\AOP_Hang_Nang_NTB_T7-T12_2026_v8_1.docx"
doc = docx.Document(docx_path)

print("=== PARAGRAPHS WITH NUMBERS OR COSTS ===")
for idx, para in enumerate(doc.paragraphs):
    text = para.text.strip()
    if text and any(k in text.lower() for k in ["chi phí", "triệu", "vnđ", "nhân sự", "định biên"]):
        print(f"[{idx}]: {text}")
