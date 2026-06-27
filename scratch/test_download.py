import urllib.request
import json
import os

url = "https://raw.githubusercontent.com/NorakGithub/cambodia-gazetteer/main/provinces/banteay_meanchey.json"
try:
    print("Downloading Banteay Meanchey JSON...")
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req) as response:
        data = json.loads(response.read().decode())
    print("Keys in province JSON:", list(data.keys()) if isinstance(data, dict) else "Not a dict")
    if isinstance(data, list) and len(data) > 0:
        print("First item in list keys:", list(data[0].keys()))
        print("First item sample:", data[0])
except Exception as e:
    print("Error:", e)
