import os
from PIL import Image
import numpy as np
import sys

sys.stdout.reconfigure(encoding='utf-8')

artifact_dir = r"C:\Users\lap4all\.gemini\antigravity-ide\brain\5935f00b-cd19-4de8-83bb-03170a8c8606"

for fname in ["media__1785944466174.png", "media__1785944607661.png"]:
    fp = os.path.join(artifact_dir, fname)
    if os.path.exists(fp):
        im = Image.open(fp).convert("RGB")
        arr = np.array(im)
        print(f"=== {fname} ===")
        print(f"Mean RGB: {arr.mean(axis=(0,1))}")
        print(f"Size: {im.size}")
