import sys, re, json, requests, base64
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.asymmetric import padding
from cryptography.hazmat.primitives.serialization import load_pem_public_key

sys.stdout.reconfigure(encoding='utf-8')

with open('scratch/app.js', 'r', encoding='utf-8') as f:
    js = f.read()

print('=== FETCH CALLS IN APP.JS ===')
fetches = re.findall(r'fetch\([^)]+\)', js)
for f in fetches:
    print('Fetch:', f)

# Authenticate session
session = requests.Session()
r_pk = session.get('https://www.quyhoachbuucuc.info/pubkey')
public_key = load_pem_public_key(r_pk.json()['pubkey'].encode('utf-8'))
encrypted = public_key.encrypt(
    'taghjnxorjvq'.encode('utf-8'),
    padding.OAEP(mgf=padding.MGF1(algorithm=hashes.SHA256()), algorithm=hashes.SHA256(), label=None)
)
epassword = base64.b64encode(encrypted).decode('utf-8')
session.post('https://www.quyhoachbuucuc.info/login', json={'user': 'ntb', 'epassword': epassword})

# Try fetching data endpoints extracted from JS
urls_to_try = [
    'https://www.quyhoachbuucuc.info/api/regions',
    'https://www.quyhoachbuucuc.info/api/ntb',
    'https://www.quyhoachbuucuc.info/api/wards',
    'https://www.quyhoachbuucuc.info/api/pos',
    'https://www.quyhoachbuucuc.info/api/data',
    'https://www.quyhoachbuucuc.info/api/summary',
    'https://www.quyhoachbuucuc.info/api/export',
    'https://www.quyhoachbuucuc.info/data.json',
    'https://www.quyhoachbuucuc.info/api/data.json'
]

for url in urls_to_try:
    r = session.get(url)
    print(f'{url} -> status {r.status_code}, len {len(r.text)}')
    if r.status_code == 200:
        with open(f'scratch/{url.split("/")[-1]}', 'w', encoding='utf-8') as out_f:
            out_f.write(r.text)
