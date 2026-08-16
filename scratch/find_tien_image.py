import os
import sys
from PIL import Image

sys.stdout.reconfigure(encoding='utf-8')

artifact_dir = r"C:\Users\lap4all\.gemini\antigravity-ide\brain\5935f00b-cd19-4de8-83bb-03170a8c8606"

files = [os.path.join(artifact_dir, f) for f in os.listdir(artifact_dir) if f.endswith(('.png', '.jpg', '.jpeg'))]
files.sort(key=lambda x: os.path.getmtime(x), reverse=True)

print("Latest image files in artifact directory:")
for f in files[:10]:
    im = Image.open(f)
    print(f"  {os.path.basename(f)}: size={im.size}, mtime={os.path.getmtime(f)}")
