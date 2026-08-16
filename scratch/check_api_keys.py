import os
from dotenv import load_dotenv

BASE_DIR = r"c:\Users\lap4all\Documents\Auto report"
load_dotenv(os.path.join(BASE_DIR, ".env"))
load_dotenv(r"c:\Users\lap4all\Desktop\New folder\.env")

print("Checking environment variables for AI keys...")
keys = [k for k in os.environ.keys() if "KEY" in k.upper() or "TOKEN" in k.upper() or "GEMINI" in k.upper() or "AI" in k.upper()]
print("Matching environment variables found:")
for k in keys:
    # Print the name and whether it has a value, but not the value itself for security
    print(f"- {k}: {'[HAS VALUE]' if os.environ[k] else '[EMPTY]'}")
