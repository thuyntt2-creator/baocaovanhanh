import os
from PIL import Image
import sys

sys.stdout.reconfigure(encoding='utf-8')

img_path = r"C:\Users\lap4all\.gemini\antigravity-ide\brain\5935f00b-cd19-4de8-83bb-03170a8c8606\web_maps\08_whatif_card_on_map.png"
if os.path.exists(img_path):
    im = Image.open(img_path)
    print(f"08_whatif_card_on_map.png -> Size: {im.size}, Size on disk: {os.path.getsize(img_path)} bytes")
else:
    print("Image not found!")
