import os
import sys
import requests
from dotenv import load_dotenv

sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = r"c:\Users\lap4all\Documents\Auto report"
load_dotenv(os.path.join(BASE_DIR, ".env"))
load_dotenv(r"c:\Users\lap4all\Desktop\New folder\.env")

api_key = os.environ.get("GEMINI_API_KEY")

def main():
    url = f"https://generativelanguage.googleapis.com/v1beta/models?key={api_key}"
    response = requests.get(url, timeout=20)
    if response.status_code == 200:
        data = response.json()
        print("Available models:")
        for m in data.get('models', []):
            print(f"- {m['name']} (Supported methods: {m.get('supportedGenerationMethods')})")
    else:
        print(f"Error {response.status_code}: {response.text}")

if __name__ == "__main__":
    main()
