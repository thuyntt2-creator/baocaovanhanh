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

def call_gemini_rest(prompt):
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key={api_key}"
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
        return f"HTTP Error {response.status_code}: {response.text}"

def main():
    print("Calling Gemini 2.0 Flash via REST...")
    res = call_gemini_rest("Xin chào! Hãy viết 1 câu ngắn chào buổi sáng bằng tiếng Việt.")
    print("Gemini response:")
    print(res)

if __name__ == "__main__":
    main()
