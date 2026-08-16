import os
import glob

old_token = "2073027751071649792:RfgqBJY4QtV18udu1wMEfhmoRNI4hgBv"
correct_token = "2067164759710552066:RfgqBJY4QtV18udu1wMEfhmoRNI4hgBv"

target_dir = r"c:\Users\lap4all\Documents\Auto report"
pattern = os.path.join(target_dir, "*.py")

files_updated = 0

for filepath in glob.glob(pattern):
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()
        
        if old_token in content:
            updated_content = content.replace(old_token, correct_token)
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(updated_content)
            print(f"Restored token prefix in: {os.path.basename(filepath)}")
            files_updated += 1
    except Exception as e:
        print(f"Error updating {os.path.basename(filepath)}: {e}")

print(f"\nDone. Restored {files_updated} files.")
