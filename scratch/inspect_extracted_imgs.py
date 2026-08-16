import os, glob, sys
from PIL import Image

sys.stdout.reconfigure(encoding='utf-8')

img_dir = r'C:\Users\lap4all\Documents\Auto report\scratch\w32_docx_images'
img_files = glob.glob(os.path.join(img_dir, "*.png"))

print(f"=== INSPECTING {len(img_files)} IMAGES ===")
for f in img_files:
    fname = os.path.basename(f)
    try:
        with Image.open(f) as img:
            w, h = img.size
            mode = img.mode
            print(f"{fname:12s}: {w}x{h}, mode={mode}, size={os.path.getsize(f)} bytes")
    except Exception as e:
        print(f"{fname:12s}: error {e}")
