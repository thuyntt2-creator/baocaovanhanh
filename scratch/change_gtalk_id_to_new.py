import os
import glob

old_id = "2067164759710552066"
new_id = "2073027751071649792"

target_dir = r"c:\Users\lap4all\Documents\Auto report"
pattern = os.path.join(target_dir, "*.py")

files_updated = 0

for filepath in glob.glob(pattern):
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()
        
        if old_id in content:
            updated_content = content.replace(old_id, new_id)
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(updated_content)
            print(f"Updated: {os.path.basename(filepath)}")
            files_updated += 1
    except Exception as e:
        print(f"Error updating {os.path.basename(filepath)}: {e}")

print(f"\nDone. Updated {files_updated} files.")
