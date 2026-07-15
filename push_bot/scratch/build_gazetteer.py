import urllib.request
import json
import os

def main():
    provinces_url = "https://raw.githubusercontent.com/NorakGithub/cambodia-gazetteer/main/provinces.json"
    print("Downloading provinces.json...")
    req = urllib.request.Request(provinces_url, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req) as response:
        provinces = json.loads(response.read().decode('utf-8'))
    
    all_communes = []
    
    for prov in provinces:
        prov_code = prov.get("code")
        prov_en = prov.get("english").replace(" Province", "").replace(" Capital", "").strip()
        prov_kh = prov.get("local").replace("ខេត្ត", "").replace("រាជធានី", "").strip()
        prov_id = prov.get("id")
        
        prov_url = f"https://raw.githubusercontent.com/NorakGithub/cambodia-gazetteer/main/provinces/{prov_id}.json"
        
        print(f"Downloading {prov_id}.json...")
        try:
            req_prov = urllib.request.Request(prov_url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req_prov) as resp_prov:
                districts = json.loads(resp_prov.read().decode('utf-8'))
            
            for dist in districts:
                # E.g., dist = {"id": "mongkol_borei", "english": "Mongkol Borei District", "local": "ស្រុកមង្គលបុរី", "communes": [...]}
                dist_en = dist.get("english").replace(" District", "").replace(" Municipality", "").replace(" Khan", "").replace(" Krong", "").strip()
                dist_kh = dist.get("local").replace("ស្រុក", "").replace("ខណ្ឌ", "").replace("ក្រុង", "").strip()
                
                for comm in dist.get("communes", []):
                    comm_en = comm.get("english").replace(" Commune", "").replace(" Sangkat", "").strip()
                    comm_kh = comm.get("local").replace("ឃុំ", "").replace("សង្កាត់", "").strip()
                    
                    all_communes.append({
                        "prov_code": prov_code,
                        "prov_en": prov_en,
                        "prov_kh": prov_kh,
                        "dist_en": dist_en,
                        "dist_kh": dist_kh,
                        "comm_en": comm_en,
                        "comm_kh": comm_kh
                    })
        except Exception as e:
            print(f"Error downloading {prov_id}: {e}")
            
    # Save the processed list of communes
    out_path = "cambodia_gazetteer.json"
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(all_communes, f, indent=2, ensure_ascii=False)
    print(f"Saved {len(all_communes)} communes to {out_path}")

if __name__ == "__main__":
    main()
