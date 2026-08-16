import glob
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

dl_dir = r"C:\Users\lap4all\Downloads"
ws_dir = r"c:\Users\lap4all\Documents\Auto report"

pdf_files = glob.glob(os.path.join(dl_dir, "*.pdf")) + glob.glob(os.path.join(ws_dir, "*.pdf"))
print("Existing PDF files:", pdf_files)

# Check if we can convert docx or md to pdf using docx2pdf or reportlab or markdown-pdf if available
try:
    from docx2pdf import convert
    print("docx2pdf is installed!")
except ImportError:
    print("docx2pdf is not installed.")
