import os

def find():
    cwd = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    for root, dirs, files in os.walk(cwd):
        if any(p in root.lower() for p in ["venv", ".git", "playwright_profile", ".gemini"]):
            continue
        for f in files:
            if f.endswith(".py"):
                path = os.path.join(root, f)
                try:
                    with open(path, "r", encoding="utf-8", errors="ignore") as file:
                        content = file.read()
                        if "baocao.ghn.vn" in content or "export_table" in content:
                            print(f"Found match: {path}")
                except Exception as e:
                    pass

if __name__ == "__main__":
    find()
