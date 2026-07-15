import json
import downloader

print("Loading config...")
with open("config.json", encoding="utf-8") as f:
    cfg = json.load(f)

print("Fetching one post office for 'BAT'...")
try:
    results = downloader.download_post_offices(cfg["api"], "BAT", limit=1)
    if results:
        with open("sample_post_office.json", "w", encoding="utf-8") as f_out:
            json.dump(results[0], f_out, indent=2, ensure_ascii=False)
        print("Successfully saved sample post office to sample_post_office.json!")
    else:
        print("No post offices found starting with BAT.")
except Exception as e:
    print(f"Error fetching: {e}")
