import sys, requests, base64, json, re, os
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
print('Login status:', r_login.status_code)
print('Login response:', r_login.text)

# 4. Fetch /web/index.html
r_app = session.get('https://www.quyhoachbuucuc.info/web/index.html')
print('App HTML length:', len(r_app.text))

# Save app HTML to inspect
with open('scratch/app.html', 'w', encoding='utf-8') as f:
    f.write(r_app.text)

# Find scripts
scripts = re.findall(r'src=["\'](.*?)["\']', r_app.text)
print('Scripts:', scripts)

# Fetch any custom JS files
for s in scripts:
    if not s.startswith('http'):
        url_s = f'https://www.quyhoachbuucuc.info/web/{s}' if not s.startswith('/') else f'https://www.quyhoachbuucuc.info{s}'
        r_js = session.get(url_s)
        print(f'JS {s} length:', len(r_js.text))
        fname = os.path.basename(s.split('?')[0])
        with open(f'scratch/{fname}', 'w', encoding='utf-8') as f:
            f.write(r_js.text)

# Try fetching API endpoints like /api/data, /data, /api/wards, /api/ntb, etc.
endpoints = ['/data', '/api/data', '/api/ntb', '/api/planning', '/api/wards', '/api/pos', '/api/all', '/config', '/summary']
for ep in endpoints:
    r_ep = session.get(f'https://www.quyhoachbuucuc.info{ep}')
    if r_ep.status_code == 200:
        print(f'FOUND ENDPOINT {ep}: length {len(r_ep.text)}')
        try:
            print(f'JSON keys for {ep}:', list(r_ep.json().keys()) if isinstance(r_ep.json(), dict) else f'List of len {len(r_ep.json())}')
        except:
            print(f'Snippet {ep}:', r_ep.text[:200])
