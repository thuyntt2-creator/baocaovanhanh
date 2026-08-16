import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

# Try using pytesseract or PIL or simple string search if plain text is embedded, or check image metadata / file list
import glob

img_dir = r'C:\Users\lap4all\Documents\Auto report\scratch\w32_docx_images'
img_files = glob.glob(os.path.join(img_dir, "*.png"))

print(f"Found {len(img_files)} images in {img_dir}")
for f in img_files:
    fname = os.path.basename(f)
    size = os.path.getsize(f)
    print(f"Image: {fname}, size: {size} bytes")
