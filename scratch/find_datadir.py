import sys, re

sys.stdout.reconfigure(encoding='utf-8')

with open('scratch/app.js', 'r', encoding='utf-8') as f:
    js = f.read()

for line in js.split('\n'):
    if 'DATA_DIR' in line:
        print(line)
