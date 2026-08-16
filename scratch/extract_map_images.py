import zipfile
import os
import shutil
import glob
import sys

sys.stdout.reconfigure(encoding='utf-8')

artifact_dir = r"C:\Users\lap4all\.gemini\antigravity-ide\brain\5935f00b-cd19-4de8-83bb-03170a8c8606"
workspace_dir = r"c:\Users\lap4all\Documents\Auto report"
doc_with_img = os.path.join(workspace_dir, "Bao_Cao_Quy_Hoach_Buu_Cuc_NTB_Theo_DVHC_Moi_Co_Hinh_AM.docx")

print("Checking DOCX with images:", doc_with_img)
if os.path.exists(doc_with_img):
    with zipfile.ZipFile(doc_with_img) as z:
        img_files = [f for f in z.namelist() if f.startswith("word/media/")]
        print(f"Found {len(img_files)} images inside DOCX!")
        for img in img_files:
            fname = os.path.basename(img)
            out_dest = os.path.join(artifact_dir, fname)
            with z.open(img) as src, open(out_dest, "wb") as dst:
                shutil.copyfileobj(src, dst)
            print(f"Extracted: {fname} -> {out_dest}")

# Also check workspace doc_img_*.png files
for png in glob.glob(os.path.join(workspace_dir, "doc_img_*.png")):
    fname = os.path.basename(png)
    out_dest = os.path.join(artifact_dir, fname)
    shutil.copy2(png, out_dest)
    print(f"Copied workspace PNG: {fname} -> {out_dest}")
