import sys
import json
import re

sys.stdout.reconfigure(encoding='utf-8')

CONFIG_PATH = 'config.json'
REGISTERED_GROUPS_PATH = 'registered_groups.json'

def load_config():
    with open(CONFIG_PATH, encoding="utf-8") as f:
        return json.load(f)

def load_registered_groups():
    with open(REGISTERED_GROUPS_PATH, encoding="utf-8") as f:
        return json.load(f)

def get_forward_mapping(cfg):
    mapping = dict(cfg["telegram"].get("forward_mapping", {}))
    
    # Auto-detect from registered group titles!
    for g in load_registered_groups():
        chat_id_str = str(g["chat_id"])
        title = g.get("title", "")
        # Look for handles like PNPP014, SVAP001, KANP001 in the title
        found_handles = re.findall(r'\b[A-Z]{3}P\d{3}\b', title.upper())
        
        # If no full handles found, try matching 3-letter branch codes (like SIH, KOH)
        if not found_handles:
            words = re.findall(r'\b[A-Z]{3}\b', title.upper())
            for w in words:
                known_prefixes = []
                for zb in cfg.get("zone_branches", {}).values():
                    known_prefixes.extend([p.strip().upper() for p in zb.split(",") if p.strip()])
                
                if w in known_prefixes:
                    found_handles.append(f"{w}P001")
        
        if found_handles:
            if chat_id_str not in mapping:
                mapping[chat_id_str] = found_handles
                
    return mapping

def get_all_forward_groups(cfg):
    mapping = get_forward_mapping(cfg)
    configured  = list(mapping.keys())
    registered  = [str(g["chat_id"]) for g in load_registered_groups()]
    all_groups = list(dict.fromkeys(configured + registered))
    zone_groups = set(str(k) for k in cfg.get("zone_forward_mapping", {}).keys())
    return [g for g in all_groups if g not in zone_groups]

cfg = load_config()
forward_groups = get_all_forward_groups(cfg)
forward_mapping = get_forward_mapping(cfg)

# Map chat ID to title
group_titles = {}
for g in load_registered_groups():
    group_titles[str(g["chat_id"])] = g.get("title", "Unknown Title")

# 1. Plain "push" regular groups
print("--- REGULAR PUSH DESTINATIONS (Plain 'push') ---")
for group_id_str in forward_groups:
    allowed = forward_mapping.get(str(group_id_str), ["*"])
    title = group_titles.get(group_id_str, "Config-only Group")
    print(f"Group: {title} ({group_id_str}) -> Handles: {', '.join(allowed)}")

# 2. "push zone" destinations
print("\n--- ZONE PUSH DESTINATIONS ('push zone') ---")
zone_fwd_map = cfg.get("zone_forward_mapping", {})
for group_id_str, zone_key in zone_fwd_map.items():
    title = group_titles.get(group_id_str, "Config-only Zone Group")
    print(f"Group: {title} ({group_id_str}) -> Zone: {zone_key.upper()}")
