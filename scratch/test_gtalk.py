import sys, os, requests, urllib3, time
from PIL import Image
sys.stdout.reconfigure(encoding='utf-8')
urllib3.disable_warnings()
from dotenv import load_dotenv

env_path = r'c:\Users\lap4all\Desktop\New folder\.env'
if os.path.exists(env_path):
    load_dotenv(dotenv_path=env_path, override=True)
    print("Loaded .env from:", env_path)

token = os.environ.get('OPR_GTALK_OA_TOKEN') or '2067164759497973760:RfgqBJY4QtV18udu1wMEfhmoRNI4hgBv'
channel = os.environ.get('OPR_GTALK_CHANNEL_ID') or '2067283005274091520'

print(f"Token: {token}")
print(f"Channel: {channel}")

img = Image.new('RGB', (100, 100), color = 'red')
img_path = 'test_gtalk.png'
img.save(img_path)

width, height = img.size
file_size = os.path.getsize(img_path)

# Step 1: Initiate Upload
initiate_url = "https://mbff.ghn.vn/api/gtalk/initiate-upload"
payload_init = {
    "ChannelId": channel,
    "FileName": os.path.basename(img_path),
    "FileSize": str(file_size),
    "MimeType": "image/png",
    "Metadata": f'{{"width": {width}, "height": {height}}}',
    "oaToken": token
}
headers = {"Content-Type": "application/json"}
res_init = requests.post(initiate_url, json=payload_init, headers=headers, timeout=20, verify=False)
print("Step 1 initiate response:", res_init.status_code, res_init.text)

if res_init.status_code == 200:
    res_data = res_init.json()
    if res_data.get("errorCode") == "success":
        presigned_url = res_data["data"]["PresignedURL"]
        upload_id = res_data["data"]["UploadId"]
        
        # Step 2: Upload S3
        with open(img_path, "rb") as f:
            res_put = requests.put(presigned_url, data=f, headers={"Content-Type": "image/png"}, timeout=60, verify=False)
        print("Step 2 S3 PUT response:", res_put.status_code, res_put.text)
        
        # Step 3: Complete Upload
        complete_url = "https://mbff.ghn.vn/api/gtalk/complete-upload"
        payload_complete = {
            "oaToken": token,
            "UploadId": upload_id
        }
        res_comp = requests.post(complete_url, json=payload_complete, headers=headers, timeout=20, verify=False)
        print("Step 3 complete upload response:", res_comp.status_code, res_comp.text)
        
        if res_comp.status_code == 200:
            res_data_comp = res_comp.json()
            if res_data_comp.get("errorCode") == "success":
                file_id = res_data_comp["data"]["Id"]
                
                # Step 4: Send Message
                send_url = "https://mbff.ghn.vn/api/gtalk/send-message"
                client_msg_id = str(int(time.time() * 1000))
                payload_send = {
                    "channelId": channel,
                    "clientMsgId": client_msg_id,
                    "content": {
                        "parseMode": "HTML",
                        "attachment": {
                            "caption": "Test GTalk image send",
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
                    "oaToken": token
                }
                res_send = requests.post(send_url, json=payload_send, headers=headers, timeout=20, verify=False)
                print("Step 4 send message response:", res_send.status_code, res_send.text)

if os.path.exists(img_path):
    os.remove(img_path)
