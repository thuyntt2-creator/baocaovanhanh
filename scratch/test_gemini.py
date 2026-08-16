import os
import sys
from dotenv import load_dotenv

sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = r"c:\Users\lap4all\Documents\Auto report"
load_dotenv(os.path.join(BASE_DIR, ".env"))
load_dotenv(r"c:\Users\lap4all\Desktop\New folder\.env")

api_key = os.environ.get("GEMINI_API_KEY")

try:
    import google.generativeai as genai
    print("google.generativeai package is installed.")
    
    genai.configure(api_key=api_key)
    
    # Let's try listing models or calling generating content
    model = genai.GenerativeModel('gemini-1.5-flash')
    response = model.generate_content("Xin chào! Hãy viết 1 câu ngắn chào buổi sáng bằng tiếng Việt.")
    print("Gemini response:")
    print(response.text.strip())
except Exception as e:
    print(f"Error calling Gemini: {e}")
