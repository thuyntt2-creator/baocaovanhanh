import os, sys

sys.stdout.reconfigure(encoding='utf-8')

docx_file1 = r'C:\Users\lap4all\Downloads\Quy_Hoach_MANG_LUOI_NTB_Co_Nha_Trang_Final.docx'
pdf_file1 = r'C:\Users\lap4all\Downloads\Quy_Hoach_MANG_LUOI_NTB_2026.pdf'

docx_file2 = r'C:\Users\lap4all\Downloads\Thu_trinh_Quy_Hoach_Mang_Luoi_NTB_2026_Co_Nha_Trang.docx'
pdf_file2 = r'C:\Users\lap4all\Downloads\Bang_Quy_Hoach_5_Tinh_NTB_2026.pdf'

try:
    import win32com.client
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    
    # Convert File 1
    if os.path.exists(docx_file1):
        doc = word.Documents.Open(docx_file1)
        doc.SaveAs(pdf_file1, FileFormat=17) # 17 = wdFormatPDF
        doc.Close()
        print(f"Successfully converted PDF 1: {pdf_file1}")
        
    # Convert File 2
    if os.path.exists(docx_file2):
        doc = word.Documents.Open(docx_file2)
        doc.SaveAs(pdf_file2, FileFormat=17) # 17 = wdFormatPDF
        doc.Close()
        print(f"Successfully converted PDF 2: {pdf_file2}")
        
    word.Quit()

except Exception as e:
    print(f"Win32com failed: {e}")
    try:
        from docx2pdf import convert
        convert(docx_file1, pdf_file1)
        convert(docx_file2, pdf_file2)
        print("Successfully converted using docx2pdf!")
    except Exception as e2:
        print(f"docx2pdf failed: {e2}")
