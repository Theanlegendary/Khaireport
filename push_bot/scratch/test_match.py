import json
import pandas as pd
import re
import sys

sys.stdout.reconfigure(encoding='utf-8')

# Load gazetteer
with open("cambodia_gazetteer.json", "r", encoding="utf-8") as f:
    gazetteer = json.load(f)

# Load pickup branches
df = pd.read_csv("pickup_branch_lookup.csv", encoding="utf-8")

# Let's clean the name and try to match
# We want to match:
# 1. Province (via Branch Code)
# 2. Commune EN / Commune Khmer to gazetteer's comm_en or comm_kh

# Let's map Branch Code to English Province name first to match gazetteer's prov_en
# We can do this based on config.json or custom map
BRANCH_TO_PROVINCE_EN = {
    "PRE": "Prey Veng",
    "PNP": "Phnom Penh",
    "SVA": "Svay Rieng",
    "KAN": "Kandal",
    "KAM": "Kampot",
    "KOH": "Koh Kong",
    "SIH": "Preah Sihanouk", # Sihanoukville
    "SPE": "Kampong Speu",
    "TAK": "Takeo",
    "BAN": "Banteay Meanchey",
    "BAT": "Battambang",
    "CHH": "Kampong Chhnang",
    "PUR": "Pursat",
    "SIE": "Siem Reap",
    "PRH": "Preah Vihear",
    "ODD": "Oddar Meanchey",
    "THO": "Kampong Thom",
    "CHA": "Kampong Cham",
    "KRA": "Kratie",
    "TBK": "Tboung Khmum",
    "ROT": "Ratanak Kiri",
    "MON": "Mondul Kiri",
    "STU": "Stung Treng",
    "KEP": "Kep",
    "PAI": "Pailin"
}

def clean_name(name):
    # Remove code prefix, e.g., "BATA001 - " or similar
    name = str(name).strip()
    name = re.sub(r'^[A-Z0-9]{3,10}\s*-\s*', '', name, flags=re.IGNORECASE)
    return name.strip()

matched_count = 0
results = []

for idx, row in df.iterrows():
    branch_code = str(row["Branch Code"]).strip().upper()
    prov_en_target = BRANCH_TO_PROVINCE_EN.get(branch_code)
    
    commune_en_raw = clean_name(row["Commune EN"])
    commune_kh_raw = clean_name(row["Commune Khmer"])
    
    # Try exact match in gazetteer
    match = None
    for item in gazetteer:
        # Check if province matches
        if prov_en_target and item["prov_en"].lower().replace(" ", "") != prov_en_target.lower().replace(" ", ""):
            continue
            
        # Match commune name
        c_en = item["comm_en"].lower().replace(" ", "").replace("-", "")
        c_kh = item["comm_kh"].lower().replace(" ", "").replace("-", "")
        
        target_en = commune_en_raw.lower().replace(" ", "").replace("-", "")
        target_kh = commune_kh_raw.lower().replace(" ", "").replace("-", "")
        
        if c_en == target_en or c_kh == target_kh or c_en in target_en or target_en in c_en:
            match = item
            break
            
    if match:
        matched_count += 1
        results.append({
            "code": row["Pickup Branch"],
            "commune": commune_en_raw,
            "matched_prov": match["prov_kh"],
            "matched_dist": match["dist_kh"],
            "matched_comm": match["comm_kh"]
        })
    else:
        results.append({
            "code": row["Pickup Branch"],
            "commune": commune_en_raw,
            "matched_prov": BRANCH_TO_PROVINCE_EN.get(branch_code, "បាត់ដំបង"),
            "matched_dist": "ឯកភ្នំ", # fallback to Battambang / Ek Phnom default or similar
            "matched_comm": commune_kh_raw
        })

print(f"Matched {matched_count} out of {len(df)}")
# Print some samples
print("\nSamples:")
for r in results[:10]:
    print(f"Code: {r['code']} | Commune: {r['commune']} => Prov: {r['matched_prov']}, Dist: {r['matched_dist']}, Comm: {r['matched_comm']}")
