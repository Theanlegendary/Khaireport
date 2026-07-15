import json
import downloader

print("Loading config...")
with open("config.json", encoding="utf-8") as f:
    cfg = json.load(f)

print("Fetching one post office for 'BAT'...")
try:
    results = downloader.download_post_offices(cfg["api"], "BAT", limit=1)
    if results:
        print("\n★ Found post office object:")
        print(json.dumps(results[0], indent=2, ensure_ascii=False))
    else:
        print("No post offices found starting with BAT.")
except Exception as e:
    print(f"Error fetching: {e}")
