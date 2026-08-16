import os
import glob

def find():
    cwd = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    files = glob.glob(os.path.join(cwd, "**/*gtc*"), recursive=True)
    for f in files:
        print(f)

if __name__ == "__main__":
    find()
