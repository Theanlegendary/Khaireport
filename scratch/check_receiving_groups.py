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
    for g in load_registered_groups():
        chat_id_str = str(g["chat_id"])
        title = g.get("title", "")
        found_handles = re.findall(r'\b[A-Z]{3}P\d{3}\b', title.upper())
        if found_handles:
            if chat_id_str not in mapping:
                mapping[chat_id_str] = found_handles
    return mapping

def get_all_forward_groups(cfg):
    mapping = get_forward_mapping(cfg)
    configured  = list(mapping.keys())
    registered  = [str(g["chat_id"]) for g in load_registered_groups()]
    all_groups = list(dict.fromkeys(configured + registered))
    
    # Exclude zone groups from regular forwarding
    zone_groups = set(str(k) for k in cfg.get("zone_forward_mapping", {}).keys())
    return [g for g in all_groups if g not in zone_groups]

cfg = load_config()
forward_groups = get_all_forward_groups(cfg)
forward_mapping = get_forward_mapping(cfg)

print("Checking groups that would receive PREP001 after fix:")
sent_count = 0
for group_id_str in forward_groups:
    allowed = forward_mapping.get(str(group_id_str), ["*"])
    wants_all = "*" in allowed
    will_receive = wants_all or "PREP001" in allowed
    
    title = "Unknown"
    for g in load_registered_groups():
        if str(g["chat_id"]) == group_id_str:
            title = g.get("title", "")
            break
            
    if will_receive:
        print(f"Group: {group_id_str} | Title: {title} | Allowed: {allowed}")
        sent_count += 1

print(f"Total groups receiving PREP001: {sent_count}")
