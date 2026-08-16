import os
import sys

def main():
    search_dirs = [
        r"C:\Users\lap4all\Documents",
        r"C:\Users\lap4all\Desktop",
        r"C:\Users\lap4all\Downloads",
        r"C:\Users\lap4all\Documents\Auto report"
    ]
    target = "tu_dong_bao_cao_nong.py"
    
    print(f"Searching for file '{target}'...")
    found_paths = []
    for s_dir in search_dirs:
        if not os.path.exists(s_dir):
            continue
        print(f"Scanning {s_dir}...")
        for root, dirs, files in os.walk(s_dir):
            # Skip virtual environments or playwright profiles to keep it fast
            if any(p in root.lower() for p in ["venv", ".git", "playwright_profile", "appdata"]):
                continue
            if target in files:
                full_path = os.path.join(root, target)
                found_paths.append(full_path)
                print(f"Found: {full_path}")
                
    if not found_paths:
        # Expand search to whole C:\Users\lap4all but excluding appdata/local/temp
        print("Not found in common folders. Expanding search to C:\\Users\\lap4all...")
        root_dir = r"C:\Users\lap4all"
        for root, dirs, files in os.walk(root_dir):
            if any(p in root.lower() for p in ["appdata", "venv", ".git", ".gemini", "playwright"]):
                continue
            if target in files:
                full_path = os.path.join(root, target)
                found_paths.append(full_path)
                print(f"Found: {full_path}")
                
    print("\nSearch complete.")
    print(f"Found {len(found_paths)} matching files.")

if __name__ == "__main__":
    main()
