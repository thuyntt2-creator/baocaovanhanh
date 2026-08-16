import sys, requests, base64, json, os
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.asymmetric import padding
from cryptography.hazmat.primitives.serialization import load_pem_public_key

sys.stdout.reconfigure(encoding='utf-8')

session = requests.Session()

# 1. Pubkey
r_pk = session.get('https://www.quyhoachbuucuc.info/pubkey')
public_key = load_pem_public_key(r_pk.json()['pubkey'].encode('utf-8'))

# 2. Encrypt password
encrypted = public_key.encrypt(
    'taghjnxorjvq'.encode('utf-8'),
    padding.OAEP(
        mgf=padding.MGF1(algorithm=hashes.SHA256()),
        algorithm=hashes.SHA256(),
        label=None
    )
)
epassword = base64.b64encode(encrypted).decode('utf-8')

# 3. Login
r_login = session.post('https://www.quyhoachbuucuc.info/login', json={'user': 'ntb', 'epassword': epassword})
print('Login status:', r_login.status_code, r_login.text)

# 4. Download files from /api/
files = [
    "hubs.json", "wards.json", "meta.json", "rezone.json",
    "optimizer.json", "wards_new.geojson", "competitors_jt.json", "ward_centroids.json"
]

out_dir = r'C:\Users\lap4all\Documents\Auto report\scratch\web_data'
os.makedirs(out_dir, exist_ok=True)

for f in files:
    url = f'https://www.quyhoachbuucuc.info/api/{f}'
    r = session.get(url)
    print(f'Downloading {f} -> status {r.status_code}, size: {len(r.content)/(1024):.1f} KB')
    if r.status_code == 200:
        file_path = os.path.join(out_dir, f)
        with open(file_path, 'wb') as out_f:
            out_f.write(r.content)
        print(f'Saved {f} to {file_path}')
