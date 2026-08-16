import requests

channel_id = "2073028426266603520"

tokens = {
    "Bot 2067164759710552066": "2067164759710552066:RfgqBJY4QtV18udu1wMEfhmoRNI4hgBv",
    "Bot 2067164759497973760": "2067164759497973760:RfgqBJY4QtV18udu1wMEfhmoRNI4hgBv"
}

for name, token in tokens.items():
    print(f"Testing {name}...")
    payload = {
        "channelId": channel_id,
        "clientMsgId": "1719876543" + str(list(tokens.keys()).index(name)),
        "content": {
            "parseMode": "HTML",
            "text": f"Chào bạn, đây là tin nhắn kiểm tra quyền kết nối từ {name}."
        },
        "oaToken": token
    }
    try:
        response = requests.post("https://mbff.ghn.vn/api/gtalk/send-message", json=payload)
        print(f"  Status: {response.status_code}")
        print(f"  Response: {response.text}")
    except Exception as e:
        print(f"  Error: {e}")
