import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

try:
    import pytesseract
    from PIL import Image
    
    artifact_dir = r"C:\Users\lap4all\.gemini\antigravity-ide\brain\5935f00b-cd19-4de8-83bb-03170a8c8606"
    
    for f in ["media__1785837319775.png", "media__1785837847917.png", "media__1785839587931.png", "media__1785896953199.png"]:
        fp = os.path.join(artifact_dir, f)
        if os.path.exists(fp):
            txt = pytesseract.image_to_string(Image.open(fp), lang='vie')
            print(f"=== {f} ===")
            print(txt[:300].strip())
            print()
except Exception as e:
    print("OCR error or pytesseract not configured:", e)
