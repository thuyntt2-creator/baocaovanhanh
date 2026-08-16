import os
from PIL import Image
import sys

sys.stdout.reconfigure(encoding='utf-8')

artifact_dir = r"C:\Users\lap4all\.gemini\antigravity-ide\brain\5935f00b-cd19-4de8-83bb-03170a8c8606"

for f in os.listdir(artifact_dir):
    if f.endswith(('.png', '.jpg', '.jpeg')):
        fp = os.path.join(artifact_dir, f)
        try:
            im = Image.open(fp)
            print(f"{f:30s} -> Size: {im.size}, Format: {im.format}")
        except Exception as e:
            print(f"{f:30s} -> Error: {e}")
