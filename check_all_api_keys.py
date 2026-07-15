import json
import downloader

with open("config.json", encoding="utf-8") as f:
    cfg = json.load(f)

print("Fetching all post offices for 'PNP' (Phnom Penh)...")
try:
    results = downloader.download_post_offices(cfg["api"], "PNP", limit=100)
    print(f"Retrieved {len(results)} post offices.")
    if results:
        # Collect all unique keys across all items
        all_keys = set()
        for item in results:
            all_keys.update(item.keys())
            if 'branch' in item and isinstance(item['branch'], dict):
                for k in item['branch'].keys():
                    all_keys.add(f"branch_{k}")
                    
        with open("pnp_keys.txt", "w", encoding="utf-8") as f_out:
            f_out.write("All unique keys in response:\n")
            f_out.write(json.dumps(sorted(list(all_keys)), indent=2) + "\n\n")
            
            interesting_keys = ['lat', 'lon', 'lng', 'address', 'location', 'coordinate', 'x', 'y', 'province', 'district', 'commune']
            found_keys = [k for k in all_keys if any(ik in k.lower() for ik in interesting_keys)]
            f_out.write(f"Matching keys: {found_keys}\n\n")
            
            for idx, item in enumerate(results[:3]):
                f_out.write(f"Item {idx+1} ({item.get('code')}):\n")
                for k in item:
                    if any(ik in k.lower() for ik in interesting_keys):
                        f_out.write(f"  {k}: {item[k]}\n")
        print("Successfully wrote output to pnp_keys.txt!")
    else:
        print("No post offices found for PNP.")
except Exception as e:
    print(f"Error: {e}")
