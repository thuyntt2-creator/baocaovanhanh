import os
from PIL import Image
import sys

sys.stdout.reconfigure(encoding='utf-8')

artifact_dir = r"C:\Users\lap4all\.gemini\antigravity-ide\brain\5935f00b-cd19-4de8-83bb-03170a8c8606"

for f in ["media__1785896953199.png", "media__1785907585318.png", "media__1785919447849.png", "media__1785920380323.png", "media__1785921390281.png"]:
    fp = os.path.join(artifact_dir, f)
    if os.path.exists(fp):
        im = Image.open(fp)
        print(f"{f}: {im.size}, aspect: {im.size[0]/im.size[1]:.2f}")
