import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

artifact_dir = r"C:\Users\lap4all\.gemini\antigravity-ide\brain\5935f00b-cd19-4de8-83bb-03170a8c8606"
web_dir = os.path.join(artifact_dir, "web_maps")

print("Files in artifact_dir:")
for f in os.listdir(artifact_dir):
    if f.endswith(('.png', '.jpg', '.jpeg')):
        fp = os.path.join(artifact_dir, f)
        print(f"  - {f} ({os.path.getsize(fp)} bytes)")

print("\nFiles in web_maps:")
if os.path.exists(web_dir):
    for f in os.listdir(web_dir):
        if f.endswith(('.png', '.jpg', '.jpeg')):
            fp = os.path.join(web_dir, f)
            print(f"  - {f} ({os.path.getsize(fp)} bytes)")
