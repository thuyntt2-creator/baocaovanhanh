import sys, re, requests, base64

sys.stdout.reconfigure(encoding='utf-8')

with open('scratch/app.js', 'r', encoding='utf-8') as f:
    js = f.read()

print('=== URLS & PATHS IN APP.JS ===')
matches = re.findall(r'["`\'](/[^"`\']+)["`\']', js)
for m in sorted(set(matches)):
    print('Path:', m)

# Let's search for function loadData or similar
lines = js.split('\n')
for i, l in enumerate(lines):
    if 'fetch' in l or 'data' in l.lower() or 'json' in l.lower() or 'api' in l.lower() or 'region' in l.lower():
        if len(l) < 150:
            print(f'L{i+1}: {l}')
