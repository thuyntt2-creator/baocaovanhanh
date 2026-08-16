import os
import sys
import time
import requests
import urllib3
from PIL import Image

urllib3.disable_warnings()

pairs = [
    ("OA 1 + Channel 1 (from .env)", "2077276776281051136:8hMHvBBU8qXKps3mLPzgKBucPLSQPg3Y", "2073929320358825984"),
    ("OA 2 + Channel 2 (from hardcoded script)", "2067164759497973760:RfgqBJY4QtV18udu1wMEfhmoRNI4hgBv", "2067283005274091520"),
    ("OA 1 + Channel 2", "2077276776281051136:8hMHvBBU8qXKps3mLPzgKBucPLSQPg3Y", "2067283005274091520"),
    ("OA 2 + Channel 1", "2067164759497973760:RfgqBJY4QtV18udu1wMEfhmoRNI4hgBv", "2073929320358825984"),
]

img_path = "test_opr_gtalk.png"

for label, token, channel_id in pairs:
    print(f"\n==========================================")
    print(f"Testing: {label}")
    print(f"Token: {token}")
    print(f"Channel: {channel_id}")
    
    file_size = os.path.getsize(img_path)
    width, height = 100, 100

    # Step 1: Initiate
    initiate_url = "https://mbff.ghn.vn/api/gtalk/initiate-upload"
    payload_init = {
        "ChannelId": channel_id,
        "FileName": os.path.basename(img_path),
        "FileSize": str(file_size),
        "MimeType": "image/png",
        "Metadata": f'{{"width": {width}, "height": {height}}}',
        "oaToken": token
    }
    headers = {"Content-Type": "application/json"}

    res_init = requests.post(initiate_url, json=payload_init, headers=headers, timeout=20, verify=False)
    if res_init.status_code != 200 or res_init.json().get("errorCode") != "success":
        print(f"❌ Step 1 failed: {res_init.text}")
        continue
    
    res_data = res_init.json()
    presigned_url = res_data["data"]["PresignedURL"]
    upload_id = res_data["data"]["UploadId"]

    # Step 2: Upload S3
    with open(img_path, "rb") as f:
        res_put = requests.put(presigned_url, data=f, headers={"Content-Type": "image/png"}, timeout=60, verify=False)
    if res_put.status_code != 200:
        print(f"❌ Step 2 failed: {res_put.status_code}")
        continue

    # Step 3: Complete Upload
    complete_url = "https://mbff.ghn.vn/api/gtalk/complete-upload"
    res_comp = requests.post(complete_url, json={"oaToken": token, "UploadId": upload_id}, headers=headers, timeout=20, verify=False)
    if res_comp.status_code != 200 or res_comp.json().get("errorCode") != "success":
        print(f"❌ Step 3 failed: {res_comp.text}")
        continue

    file_id = res_comp.json()["data"]["Id"]

    # Step 4: Send Message
    send_url = "https://mbff.ghn.vn/api/gtalk/send-message"
    client_msg_id = str(int(time.time() * 1000))
    payload_send = {
        "channelId": channel_id,
        "clientMsgId": client_msg_id,
        "content": {
            "parseMode": "HTML",
            "attachment": {
                "caption": "Test message",
                "items": [{"image": {"fileId": file_id, "width": width, "height": height}}]
            }
        },
        "oaToken": token
    }
    res_send = requests.post(send_url, json=payload_send, headers=headers, timeout=20, verify=False)
    print(f"Step 4 Status: {res_send.status_code}")
    print(f"Step 4 Response: {res_send.text}")
