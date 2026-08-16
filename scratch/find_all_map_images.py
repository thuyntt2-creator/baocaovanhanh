import os
import glob
import sys

sys.stdout.reconfigure(encoding='utf-8')

artifact_dir = r"C:\Users\lap4all\.gemini\antigravity-ide\brain\5935f00b-cd19-4de8-83bb-03170a8c8606"
print("Artifact dir files:")
for f in os.listdir(artifact_dir):
    full_p = os.path.join(artifact_dir, f)
    if os.path.isfile(full_p):
        print(f"  {f} ({os.path.getsize(full_p)} bytes)")

dl_dir = r"C:\Users\lap4all\Downloads"
print("\nDownloads image/media files:")
for f in glob.glob(os.path.join(dl_dir, "*.*")):
    if f.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp')):
        print(f"  {os.path.basename(f)} ({os.path.getsize(f)} bytes)")
