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
    url = f"https://generativelanguage.googleapis.com/v1/models/gemini-1.5-flash:generateContent?key={api_key}"
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
        # Fallback to gemini-1.5-flash-latest on v1beta if v1 fails
        url_beta = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key={api_key}"
        response_beta = requests.post(url_beta, headers=headers, json=payload, timeout=20)
        if response_beta.status_code == 200:
            data = response_beta.json()
            try:
                return data['candidates'][0]['content']['parts'][0]['text']
            except (KeyError, IndexError):
                return "Error parsing beta response structure"
        return f"HTTP Error {response_beta.status_code}: {response_beta.text}"

def main():
    print("Calling Gemini API via REST (v1/v1beta)...")
    res = call_gemini_rest("Xin chào! Hãy viết 1 câu ngắn chào buổi sáng bằng tiếng Việt.")
    print("Gemini response:")
    print(res)

if __name__ == "__main__":
    main()
