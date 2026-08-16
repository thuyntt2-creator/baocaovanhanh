import os
import sys
import time
import requests
import urllib3
from PIL import Image

urllib3.disable_warnings()

token = "2067164759497973760:RfgqBJY4QtV18udu1wMEfhmoRNI4hgBv"
channel_id = "2067283005274091520"

img_path = "test_opr_gtalk.png"
file_size = os.path.getsize(img_path)

payload_init = {
    "ChannelId": channel_id,
    "FileName": "test.png",
    "FileSize": str(file_size),
    "MimeType": "image/png",
    "Metadata": '{"width": 100, "height": 100}',
    "oaToken": token
}
headers = {"Content-Type": "application/json"}
res_init = requests.post("https://mbff.ghn.vn/api/gtalk/initiate-upload", json=payload_init, headers=headers, verify=False)
res_data = res_init.json()
print("Initiate response:", res_data)

presigned_url = res_data["data"]["PresignedURL"]
upload_id = res_data["data"]["UploadId"]

with open(img_path, "rb") as f:
    requests.put(presigned_url, data=f, headers={"Content-Type": "image/png"}, verify=False)

res_comp = requests.post("https://mbff.ghn.vn/api/gtalk/complete-upload", json={"oaToken": token, "UploadId": upload_id}, headers=headers, verify=False)
file_id = res_comp.json()["data"]["Id"]
print("Complete response file_id:", file_id)

msg1_tele = "📊 <b>BẢNG 1: %OPR TTS THEO AM - KPI 80%</b>\n⏱️ <b>Mốc:</b> 07:00 ngày 05/08/2026"
msg1_gtalk = msg1_tele + '\n\nChi tiết đơn lỗi theo AM (<a href="https://docs.google.com/spreadsheets/d/1d3Yeu-5mBE8w5i89_dyJ0ICl1GNP7WgZrH1oQfc5j0s/edit?gid=0#gid=0"><b>xem chi tiết</b></a>)'

payload_send = {
    "channelId": channel_id,
    "clientMsgId": str(int(time.time() * 1000)),
    "content": {
        "parseMode": "HTML",
        "attachment": {
            "caption": msg1_gtalk,
            "items": [{"image": {"fileId": file_id, "width": 100, "height": 100}}]
        }
    },
    "oaToken": token
}
res_send = requests.post("https://mbff.ghn.vn/api/gtalk/send-message", json=payload_send, headers=headers, verify=False)
print("Send Response:", res_send.text)
