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

def rewrite_questions_gemini(questions_list):
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={api_key}"
    headers = {"Content-Type": "application/json"}
    
    prompt = f"""
Bạn là một quản lý vận hành giàu kinh nghiệm. Hãy viết lại (rewrite) các câu hỏi giải trình vận hành gửi các AM (Area Manager) dưới đây để chúng tự nhiên, đa dạng cấu trúc ngữ pháp và tránh trùng lặp máy móc.

Yêu cầu cụ thể:
1. Giữ nguyên 100% các từ khóa tên bưu cục (ví dụ: bưu cục (DNO) Quảng Tín, (LDO) Di Linh), các thẻ tag tên AM (ví dụ: @Trần Văn Phước, @Thái Thị Thanh Thư) và các số liệu thống kê trong câu.
2. Thay đổi cách đặt câu hỏi, sử dụng nhiều mẫu câu khác nhau (hỏi thăm sự cố, đề xuất rà soát, yêu cầu giải pháp, nhắc nhở deadline...).
3. Văn phong chuyên nghiệp, thân thiện nhưng quyết liệt, đúng chất chat công việc (Zalo/GTalk).
4. Không thêm lời chào chung chung ngoài nội dung câu hỏi.

Đầu vào là một chuỗi JSON như sau:
{json.dumps(questions_list, ensure_ascii=False, indent=2)}

Hãy trả về kết quả định dạng JSON duy nhất khớp với cấu trúc trên:
[
  {{
    "id": 1,
    "text": "Câu hỏi đã được viết lại tự nhiên và sinh động hơn"
  }}
]
    """
    
    payload = {
        "contents": [{
            "parts": [{
                "text": prompt
            }]
        }],
        "generationConfig": {
            "responseMimeType": "application/json"
        }
    }
    
    response = requests.post(url, headers=headers, json=payload, timeout=30)
    if response.status_code == 200:
        try:
            res_data = response.json()
            text_out = res_data['candidates'][0]['content']['parts'][0]['text']
            return json.loads(text_out)
        except Exception as e:
            return f"Error parsing response: {e}. Output was: {response.text}"
    else:
        return f"API Error {response.status_code}: {response.text}"

def main():
    sample_questions = [
        {"id": 1, "text": "Anh @Trần Văn Phước xem giúp bưu cục (DNO) Quảng Tín hiện đang tồn đọng 71 đơn hàng bị aging trên 5 ngày (trong đó có 8 đơn tồn trên 15 ngày). Anh chỉ đạo bưu cục rà soát kỹ từng đơn để phản hồi nguyên nhân giao/trả chậm và giải phóng hàng tồn nhanh giúp em nhé."},
        {"id": 2, "text": "Anh @Trần Văn Phước xem giúp bưu cục (DNO) Kiến Đức hiện đang tồn đọng 45 đơn hàng bị aging trên 5 ngày. Anh chỉ đạo bưu cục rà soát kỹ từng đơn để phản hồi nguyên nhân giao/trả chậm và giải phóng hàng tồn nhanh giúp em nhé."},
        {"id": 3, "text": "Chị @Thái Thị Thanh Thư xem giúp bưu cục (KHO) Cam Linh hiện đang tồn đọng 421 đơn hàng bị aging trên 5 ngày. Chị chỉ đạo bưu cục rà soát kỹ từng đơn để phản hồi nguyên nhân giao/trả chậm và giải phóng hàng tồn nhanh giúp em nhé."},
        {"id": 4, "text": "Chị @Trần Thị Nhung check giúp bưu cục (DNO) Đức Lập hiện có 75 đơn luân chuyển giao/trả bị treo trên 36h chưa đóng kiện hoặc chưa trung chuyển đi. Chị nhắc bưu cục xử lý dứt điểm các đơn treo này trong hôm nay nha."}
    ]
    
    print("Sending sample questions to Gemini for rewriting...")
    rewritten = rewrite_questions_gemini(sample_questions)
    print("\nRewritten output:")
    print(json.dumps(rewritten, ensure_ascii=False, indent=2))

if __name__ == "__main__":
    main()
