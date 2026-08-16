import os
import sys
import requests
import json
from dotenv import load_dotenv

sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = r"c:\Users\lap4all\Documents\Auto report"
load_dotenv(os.path.join(BASE_DIR, ".env"))
load_dotenv(r"c:\Users\lap4all\Desktop\New folder\.env")

api_key = os.environ.get("GEMINI_API_KEY")

def call_gemini_rest(model_name, prompt):
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent?key={api_key}"
    headers = {"Content-Type": "application/json"}
    payload = {
        "contents": [{
            "parts": [{
                "text": prompt
            }]
        }]
    }
    
    response = requests.post(url, headers=headers, json=payload, timeout=20)
    if response.status_code == 200:
        data = response.json()
        try:
            return data['candidates'][0]['content']['parts'][0]['text']
        except (KeyError, IndexError):
            return "Error parsing response structure"
    else:
        return f"Error {response.status_code}: {response.text}"

def main():
    for model in ["gemini-2.5-flash", "gemini-3.5-flash"]:
        print(f"Testing model: {model}...")
        res = call_gemini_rest(model, "Xin chào! Hãy viết 1 câu ngắn chào buổi sáng bằng tiếng Việt.")
        print(f"Response from {model}:")
        print(res)
        print("-" * 40)

if __name__ == "__main__":
    main()
