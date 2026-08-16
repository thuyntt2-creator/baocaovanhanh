import os
import glob

# For all files EXCEPT calculate_report_aging.py, we restore the original channel 2067164759710552066
old_channel = "2073027751071649792"
original_channel = "2067164759710552066"

target_dir = r"c:\Users\lap4all\Documents\Auto report"
pattern = os.path.join(target_dir, "*.py")

files_updated = 0

for filepath in glob.glob(pattern):
    filename = os.path.basename(filepath)
    if filename == "calculate_report_aging.py":
        continue
        
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()
        
        if old_channel in content:
            updated_content = content.replace(old_channel, original_channel)
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(updated_content)
            print(f"Reverted channel ID in: {filename}")
            files_updated += 1
    except Exception as e:
        print(f"Error reverting {filename}: {e}")

print(f"\nDone. Reverted {files_updated} files.")
