import os
import sys
import time
import requests
import urllib3
from dotenv import load_dotenv
from PIL import Image

urllib3.disable_warnings()

env_path = r"c:\Users\lap4all\Desktop\New folder\.env"
if os.path.exists(env_path):
    load_dotenv(dotenv_path=env_path, override=True)
else:
    load_dotenv()

# From user's script:
TELEGRAM_TOKEN = "8570130113:AAGXRiUaKBknVpgtm1_i9ZA47JRjAXmB21M"
TELEGRAM_CHAT_ID = "-5058464865"
GTALK_OA_TOKEN = "2067164759497973760:RfgqBJY4QtV18udu1wMEfhmoRNI4hgBv"
GTALK_CHANNEL_ID = "2067283005274091520"

# Override config if defined in env
GTALK_OA_TOKEN = os.environ.get("OPR_GTALK_OA_TOKEN") or os.environ.get("GTALK_OA_TOKEN") or GTALK_OA_TOKEN
GTALK_CHANNEL_ID = os.environ.get("OPR_GTALK_CHANNEL_ID") or os.environ.get("GTALK_CHANNEL_ID") or GTALK_CHANNEL_ID

print(f"Token: {GTALK_OA_TOKEN}")
print(f"Channel ID: {GTALK_CHANNEL_ID}")

# Test image creation
img_path = "test_opr_gtalk.png"
img = Image.new("RGB", (100, 100), color="red")
img.save(img_path)

file_size = os.path.getsize(img_path)
width, height = img.size

# Step 1: Initiate Upload
initiate_url = "https://mbff.ghn.vn/api/gtalk/initiate-upload"
payload_init = {
    "ChannelId": GTALK_CHANNEL_ID,
    "FileName": os.path.basename(img_path),
    "FileSize": str(file_size),
    "MimeType": "image/png",
    "Metadata": f'{{"width": {width}, "height": {height}}}',
    "oaToken": GTALK_OA_TOKEN
}
headers = {"Content-Type": "application/json"}

print("\n--- STEP 1: INITIATE UPLOAD ---")
res_init = requests.post(initiate_url, json=payload_init, headers=headers, timeout=20, verify=False)
print("Status:", res_init.status_code)
print("Response:", res_init.text)
if res_init.status_code != 200:
    sys.exit(1)

res_data = res_init.json()
if res_data.get("errorCode") != "success":
    print("Initiate failed errorCode:", res_data.get("errorCode"), "error:", res_data.get("error"))
    sys.exit(1)

presigned_url = res_data["data"]["PresignedURL"]
upload_id = res_data["data"]["UploadId"]
print("PresignedURL obtained, UploadId:", upload_id)

# Step 2: Upload to S3
print("\n--- STEP 2: UPLOAD TO S3 ---")
with open(img_path, "rb") as f:
    headers_put = {"Content-Type": "image/png"}
    res_put = requests.put(presigned_url, data=f, headers=headers_put, timeout=60, verify=False)
print("PUT Status:", res_put.status_code)
if res_put.status_code != 200:
    print("PUT Response:", res_put.text)
    sys.exit(1)

# Step 3: Complete Upload
print("\n--- STEP 3: COMPLETE UPLOAD ---")
complete_url = "https://mbff.ghn.vn/api/gtalk/complete-upload"
payload_complete = {
    "oaToken": GTALK_OA_TOKEN,
    "UploadId": upload_id
}
res_comp = requests.post(complete_url, json=payload_complete, headers=headers, timeout=20, verify=False)
print("Complete Status:", res_comp.status_code)
print("Response:", res_comp.text)
if res_comp.status_code != 200:
    sys.exit(1)

res_data_comp = res_comp.json()
if res_data_comp.get("errorCode") != "success":
    print("Complete failed error:", res_data_comp.get("error"))
    sys.exit(1)

file_id = res_data_comp["data"]["Id"]
print("File ID:", file_id)

# Step 4: Send Message
print("\n--- STEP 4: SEND MESSAGE ---")
send_url = "https://mbff.ghn.vn/api/gtalk/send-message"
client_msg_id = str(int(time.time() * 1000))
caption = "Test OPR GTalk message"
payload_send = {
    "channelId": GTALK_CHANNEL_ID,
    "clientMsgId": client_msg_id,
    "content": {
        "parseMode": "HTML",
        "attachment": {
            "caption": caption,
            "items": [
                {
                    "image": {
                        "fileId": file_id,
                        "width": width,
                        "height": height
                    }
                }
            ]
        }
    },
    "oaToken": GTALK_OA_TOKEN
}
res_send = requests.post(send_url, json=payload_send, headers=headers, timeout=20, verify=False)
print("Send Status:", res_send.status_code)
print("Send Response:", res_send.text)
