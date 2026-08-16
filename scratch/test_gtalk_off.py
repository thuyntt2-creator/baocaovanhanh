import requests
import json

url = "https://mbff.ghn.vn/api/gtalk/send-message"
token = "2067164759497973760:RfgqBJY4QtV18udu1wMEfhmoRNI4hgBv"
group_id = "2067164759710552066"

payload = {
    "channelId": str(group_id),
    "clientMsgId": "1722234567890",
    "content": {
        "parseMode": "HTML",
        "text": "🧪 [TEST] Kiểm tra kết nối GTalk API cho thông báo Off Tuyến."
    },
    "oaToken": token
}

try:
    r = requests.post(url, json=payload, timeout=15, verify=False)
    print("Status:", r.status_code)
    print("Response:", r.text)
except Exception as e:
    print("Error:", e)
