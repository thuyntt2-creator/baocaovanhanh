import requests
import re
import sys

sys.stdout.reconfigure(encoding='utf-8')

url = "https://www.quyhoachbuucuc.info/web/app.js?v=1785502319"
resp = requests.get(url)
code = resp.text

print(f"app.js size: {len(code)} bytes")

# Search for maplibre / map initialization
matches = re.findall(r'new\s+maplibregl\.Map\s*\(\s*\{[^}]*\}', code)
print("maplibregl.Map instantiations:", matches)

# Search for variable names assigned to maplibregl.Map
var_matches = re.findall(r'(\w+)\s*=\s*new\s+maplibregl\.Map', code)
print("Variables assigned new maplibregl.Map:", var_matches)

# Search for flyTo or setCenter calls in app.js
fly_matches = re.findall(r'\w+\.flyTo\(', code)
print("flyTo calls:", set(fly_matches))

# Search for what-if or popup panel triggers in app.js
popup_matches = re.findall(r'what-if|\.show|\.render|modal|drawer|panel|prop-|rezon', code, re.IGNORECASE)
print("Found keywords count:", len(popup_matches))
