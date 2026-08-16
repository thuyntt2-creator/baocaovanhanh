import os
import glob

downloads_dir = r"C:\Users\lap4all\Downloads"
workspace_dir = r"c:\Users\lap4all\Documents\Auto report"

print("Searching in Downloads:")
pattern_dl = os.path.join(downloads_dir, "*AOP_Hang_Nang*")
for f in glob.glob(pattern_dl):
    print(f)

print("Searching for AOP files in Downloads:")
for f in glob.glob(os.path.join(downloads_dir, "*AOP*")):
    print(f)

print("Searching in Workspace:")
for root, dirs, files in os.walk(workspace_dir):
    for f in files:
        if "AOP" in f or "Hang_Nang" in f:
            print(os.path.join(root, f))
