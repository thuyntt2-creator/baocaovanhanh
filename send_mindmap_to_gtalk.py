import os
import sys
import json
import requests
from datetime import datetime
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Load env variables
env_paths = [
    os.path.join(BASE_DIR, ".env"),
    r"c:\Users\lap4all\Desktop\New folder\.env"
]
for p in env_paths:
    if os.path.exists(p):
        load_dotenv(dotenv_path=p, override=True)
        break

GTALK_OA_TOKEN = os.environ.get("GTALK_OA_TOKEN") or "2067164759710552066:RfgqBJY4QtV18udu1wMEfhmoRNI4hgBv"
GTALK_CHANNEL_ID = os.environ.get("MORNING_QUESTIONS_GTALK_CHANNEL_ID") or os.environ.get("GTALK_CHANNEL_ID") or "2067164759710552066"

def main():
    print("🎨 Preparing mind map HTML template...")
    html_content = """<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background-color: #0f172a;
            color: #f8fafc;
            padding: 30px;
            margin: 0;
            display: flex;
            flex-direction: column;
            align-items: center;
        }
        h2 {
            color: #38bdf8;
            margin-bottom: 20px;
            font-size: 28px;
            text-align: center;
            font-weight: 700;
        }
        #mermaid-container {
            background-color: #1e293b;
            padding: 25px;
            border-radius: 12px;
            border: 1px solid #334155;
            box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.3);
            width: 1100px;
            display: flex;
            justify-content: center;
        }
    </style>
    <script type="module">
        import mermaid from 'https://cdn.jsdelivr.net/npm/mermaid@10/dist/mermaid.esm.min.mjs';
        mermaid.initialize({ 
            startOnLoad: true,
            theme: 'dark',
            themeVariables: {
                background: '#1e293b',
                primaryColor: '#db2777',
                primaryTextColor: '#fff',
                lineColor: '#64748b'
            }
        });
    </script>
</head>
<body>
    <h2>🗺️ SƠ ĐỒ ĐIỂM NÓNG VẬN HÀNH YẾU KÉM NTB</h2>
    <div id="mermaid-container">
        <pre class="mermaid">
graph LR
    classDef rootStyle fill:#0284c7,stroke:#0369a1,stroke-width:3px,color:#fff,font-weight:bold;
    classDef amStyle fill:#db2777,stroke:#be185d,stroke-width:2px,color:#fff,font-weight:bold;
    classDef poStyle fill:#d97706,stroke:#b45309,stroke-width:2px,color:#fff,font-weight:bold;
    classDef kpiStyle fill:#ef4444,stroke:#dc2626,stroke-width:1px,color:#fff;
    classDef warningStyle fill:#9333ea,stroke:#7e22ce,stroke-width:1px,color:#fff,font-weight:bold;

    Root["BẢN ĐỒ ĐIỂM NÓNG<br>NTB"]:::rootStyle
    
    Root --> AM1["AM Thái Thị Thanh Thư"]:::amStyle
    AM1 --> PO1["(KHO) Cam Linh"]:::poStyle
    PO1 --> KPI1["Aging: 421 đơn > 5 ngày (9 đơn > 15 ngày)"]:::kpiStyle
    PO1 --> KPI2["Treo LC: 249 đơn > 36h"]:::kpiStyle
    PO1 --> KPI3["Rớt LC: TTS 18.18%, Shopee 34.38%"]:::kpiStyle

    Root --> AM2["AM Trầm Hữu Tiến"]:::amStyle
    AM2 --> PO2["(LDO) Di Linh"]:::poStyle
    PO2 --> KPI4["Aging: 137 đơn > 5 ngày"]:::kpiStyle
    PO2 --> KPI5["Treo LC: 243 đơn > 36h"]:::kpiStyle
    AM2 --> PO3["(LDO) Đức Trọng 1"]:::poStyle
    PO3 --> KPI6["Rớt LC: Shopee 100%"]:::kpiStyle

    Root --> AM3["AM Trần Văn Phước"]:::amStyle
    AM3 --> PO4["(DNO) Quảng Tín"]:::poStyle
    PO4 --> KPI7["Aging: 71 đơn > 5 ngày (8 đơn > 15 ngày)"]:::kpiStyle
    PO4 --> KPI8["GTC: 35.08% / Tồn: 630 đơn"]:::kpiStyle
    AM3 --> PO5["(DNO) Kiến Đức"]:::poStyle
    PO5 --> KPI9["Aging: 45 đơn > 5 ngày"]:::kpiStyle
    PO5 --> KPI10["GTC: 30.14% / Tồn: 538 đơn"]:::kpiStyle
    PO5 --> KPI11["Rớt LC: Shopee 18.75%"]:::kpiStyle

    Root --> AM4["AM Hồng Bích Nga"]:::amStyle
    AM4 --> PO6["(LDO) B'Lao"]:::poStyle
    PO6 --> KPI12["Treo LC: 598 đơn > 36h"]:::warningStyle
    PO6 --> KPI13["Rớt LC: TTS 100%, Shopee 100%"]:::warningStyle

    Root --> AM5["AM Phạm Bá Thành Công"]:::amStyle
    AM5 --> PO7["(KHO) Bắc Nha Trang"]:::poStyle
    PO7 --> KPI14["Treo LC: 119 đơn > 36h"]:::kpiStyle
    AM5 --> PO8["(KHO) Vạn Ninh"]:::poStyle
    PO8 --> KPI15["Treo LC: 84 đơn > 36h"]:::kpiStyle

    Root --> AM6["AM Huỳnh Thị Kim Chi"]:::amStyle
    AM6 --> PO9["(LDO) Tân Hà Lâm Hà"]:::poStyle
    PO9 --> KPI16["Aging: 189 đơn > 5 ngày"]:::kpiStyle

    Root --> AM7["AM Nguyễn Duy Long"]:::amStyle
    AM7 --> PO10["(NTH) Thuận Nam"]:::poStyle
    PO10 --> KPI17["Rớt LC: TTS 53.16%, Shopee 28.57%"]:::kpiStyle
        </pre>
    </div>
</body>
</html>
"""

    temp_html_path = os.path.join(BASE_DIR, "temp_mindmap.html")
    with open(temp_html_path, "w", encoding="utf-8") as f:
        f.write(html_content)
        
    image_path = os.path.join(BASE_DIR, "ntb_worst_hubs_mindmap.png")
    
    print("📸 Rendering mind map with Playwright...")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        # Give enough page width and height for wide diagrams
        page = browser.new_page(viewport={"width": 1400, "height": 1300})
        page.goto(f"file:///{temp_html_path.replace('\\', '/')}")
        
        # Wait for mermaid to draw the SVG element
        page.wait_for_selector(".mermaid svg", timeout=15000)
        page.wait_for_timeout(1500) # small extra wait to ensure transition ends
        
        # Capture only the container box
        page.locator("#mermaid-container").screenshot(path=image_path)
        print(f"✔️ Screenshot captured at: {image_path}")
        browser.close()
        
    try:
        os.remove(temp_html_path)
    except:
        pass
        
    # Send to GTalk
    print("📡 Uploading image to GTalk...")
    file_name = os.path.basename(image_path)
    file_size = os.path.getsize(image_path)
    
    with open(image_path, 'rb') as f:
        file_bytes = f.read()
        
    init_payload = {
        "ChannelId": GTALK_CHANNEL_ID,
        "FileName": file_name,
        "FileSize": str(file_size),
        "MimeType": "image/png",
        "Metadata": json.dumps({"width": 1100, "height": 800}),
        "oaToken": GTALK_OA_TOKEN
    }
    
    try:
        resp_init = requests.post("https://mbff.ghn.vn/api/gtalk/initiate-upload", json=init_payload, timeout=20)
        if resp_init.status_code == 200:
            init_data = resp_init.json()
            if init_data.get("errorCode") == "success":
                presigned_url = init_data["data"]["PresignedURL"]
                upload_id = init_data["data"]["UploadId"]
                
                resp_put = requests.put(presigned_url, data=file_bytes, headers={"Content-Type": "image/png"}, timeout=40)
                if resp_put.status_code == 200:
                    resp_comp = requests.post(
                        "https://mbff.ghn.vn/api/gtalk/complete-upload", 
                        json={"oaToken": GTALK_OA_TOKEN, "UploadId": upload_id},
                        timeout=20
                    )
                    if resp_comp.status_code == 200:
                        comp_data = resp_comp.json()
                        if comp_data.get("errorCode") == "success":
                            img_id = comp_data["data"]["Id"]
                            print(f"✅ Image upload successful! Image ID: {img_id}")
                            
                            # Send message with image attachment
                            caption = "🗺️ <b>BẢN ĐỒ TƯ DUY - CÁC BƯU CỤC VẬN HÀNH YẾU KÉM NHẤT NTB</b>"
                            send_payload = {
                                "channelId": GTALK_CHANNEL_ID,
                                "clientMsgId": str(int(datetime.now().timestamp() * 1000)),
                                "content": {
                                    "parseMode": "HTML",
                                    "attachment": {
                                        "caption": caption,
                                        "items": [
                                            {"image": {"fileId": img_id, "width": 1100, "height": 800}}
                                        ]
                                    }
                                },
                                "oaToken": GTALK_OA_TOKEN
                            }
                            
                            r_send = requests.post("https://mbff.ghn.vn/api/gtalk/send-message", json=send_payload, timeout=20)
                            if r_send.status_code == 200 and r_send.json().get("errorCode") == "success":
                                print("🎉 Successfully sent mind map image to GTalk!")
                                # Also copy image to brain directory for user
                                brain_dir = r"C:\Users\lap4all\.gemini\antigravity-ide\brain\cd8bf44c-b8bf-4b11-952a-d6a745f837cd"
                                if os.path.exists(brain_dir):
                                    import shutil
                                    shutil.copy2(image_path, os.path.join(brain_dir, "ntb_worst_hubs_mindmap.png"))
                                    print("📋 Saved copy of image to brain directory.")
                                sys.exit(0)
                            else:
                                print(f"❌ Failed to send GTalk message: {r_send.text}")
        print("❌ Initiate/Complete upload failed.")
    except Exception as e:
        print(f"❌ Error broadcasting to GTalk: {e}")
        
    try:
        os.remove(image_path)
    except:
        pass

if __name__ == "__main__":
    main()
