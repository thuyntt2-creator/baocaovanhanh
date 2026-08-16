import os
from PIL import Image
import sys

sys.stdout.reconfigure(encoding='utf-8')

artifact_dir = r"C:\Users\lap4all\.gemini\antigravity-ide\brain\5935f00b-cd19-4de8-83bb-03170a8c8606"

for fname in ["media__1785944466174.png", "media__1785944607661.png"]:
    fp = os.path.join(artifact_dir, fname)
    if os.path.exists(fp):
        im = Image.open(fp)
        w, h = im.size
        # Print info about corners and center
        print(f"File: {fname} ({w}x{h})")
        # Crop 4 corners
        top_left = im.crop((0, 0, int(w*0.3), int(h*0.3)))
        center = im.crop((int(w*0.3), int(h*0.3), int(w*0.7), int(h*0.7)))
        top_right = im.crop((int(w*0.7), 0, w, int(h*0.3)))
        print("  Successfully inspected crops!")
