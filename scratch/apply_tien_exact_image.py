import os
import shutil
from PIL import Image
import sys

sys.stdout.reconfigure(encoding='utf-8')

artifact_dir = r"C:\Users\lap4all\.gemini\antigravity-ide\brain\5935f00b-cd19-4de8-83bb-03170a8c8606"
src_tien_img = os.path.join(artifact_dir, "media__1785946591077.png")

downloads_dir = r"C:\Users\lap4all\Downloads"
out_img_path = os.path.join(downloads_dir, "Di_Linh_Official_Web_Drawn_Map.png")
artifact_out_path = os.path.join(artifact_dir, "di_linh_official_web_drawn_map.png")

if not os.path.exists(src_tien_img):
    print(f"Source image not found: {src_tien_img}")
    sys.exit(1)

# Copy Tiến's exact image to destination paths
shutil.copy(src_tien_img, out_img_path)
shutil.copy(src_tien_img, artifact_out_path)

print(f"Successfully applied Tiến's exact drawing for Di Linh at:\n  - {out_img_path}\n  - {artifact_out_path}")
