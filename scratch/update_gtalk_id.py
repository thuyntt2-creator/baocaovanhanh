import os
import sys
import glob

# Ensure stdout handles encoding properly
try:
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass

# We are reverting the ID from the new one back to the old one
old_id = "2067164759497973760"
new_id = "2067164759710552066"

target_dir = r"c:\Users\lap4all\Documents\Auto report"
pattern = os.path.join(target_dir, "*.py")

files_updated = 0

for filepath in glob.glob(pattern):
    filename = os.path.basename(filepath)
    # DO NOT revert calculate_and_render_report.py
    if filename == "calculate_and_render_report.py":
        continue
        
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()
        
        if old_id in content:
            updated_content = content.replace(old_id, new_id)
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(updated_content)
            print(f"Reverted: {filename}")
            files_updated += 1
    except Exception as e:
        print(f"Error reverting {filename}: {e}")

print(f"\nDone. Reverted {files_updated} files.")
