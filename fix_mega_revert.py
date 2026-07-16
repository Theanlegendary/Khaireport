"""Revert mega pivot to only use CURRENT POST OFFICE (col 15) for MEGA/DVC/HUB."""
import os, re

HERE = os.path.dirname(os.path.abspath(__file__))
path = os.path.join(HERE, 'pivot.py')

with open(path, 'r', encoding='utf-8') as f:
    content = f.read()

# Replace build_mega_pivot
old_start = content.find('def build_mega_pivot(rows, pivot_cfg, zone_cfg):')
old_end = content.find('\n\ndef export_mega_pivot')

new_func = '''def build_mega_pivot(rows, pivot_cfg, zone_cfg):
    """
    Builds a pivot tree for MEGA check:
      Only counts orders where CURRENT POST OFFICE (col 15) contains MEGA, DVC, or HUB.
      These are parcels actually stuck at hub right now.
    """
    exclude_test = pivot_cfg.get("exclude_test", False)
    test_keywords = pivot_cfg.get("test_keywords", ["test"])
    
    exclude_statuses = {"410", "201", "520", "99", "100", "-99"}

    tree = defaultdict(lambda: defaultdict(int))
    urgent_tree = defaultdict(int)
    day_keys_seen = set()
    today = datetime.now().date()

    for row in rows:
        if not row or row[COL_ORDER_ID] in (None, ""):
            continue
        status_code = _status_code(row[COL_CURRENT_STATUS])
        if status_code in exclude_statuses:
            continue
        if exclude_test and _is_test_row(row, test_keywords):
            continue

        po = str(row[COL_CURRENT_PO] or "").strip()
        po_upper = po.upper()
        
        # Only CURRENT POST OFFICE containing MEGA, DVC, or HUB
        if not ("MEGA" in po_upper or "HUB" in po_upper or "DVC" in po_upper):
            continue

        month, day = _parse_day(row[COL_CREATED_DATE])
        if day is None:
            continue

        key = (month, day)
        tree[po][key] += 1
        day_keys_seen.add(key)

        # Check urgent: created > 1 day ago
        created_val = row[COL_CREATED_DATE]
        created_date = None
        if isinstance(created_val, datetime):
            created_date = created_val.date()
        elif created_val:
            s = str(created_val).strip().split(" ")[0]
            for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%m/%d/%Y", "%d-%m-%Y"):
                try:
                    created_date = datetime.strptime(s, fmt).date()
                    break
                except ValueError:
                    continue
        if created_date and (today - created_date).days > 1:
            urgent_tree[po] += 1

    day_keys = sorted(day_keys_seen)
    return tree, day_keys, urgent_tree'''

content = content[:old_start] + new_func + content[old_end:]

with open(path, 'w', encoding='utf-8') as f:
    f.write(content)
print("DONE pivot.py")

# Also fix mega_detail.py
path2 = os.path.join(HERE, 'mega_detail.py')
with open(path2, 'r', encoding='utf-8') as f:
    content2 = f.read()

# Replace the filter logic back to simple CURRENT POST OFFICE check
old_filter = '''    transit_statuses = {"210", "300", "302", "306", "309", "310", "311", "500"}
    seen_orders = set()

    filtered = []
    for row in data_rows:
        if not row or len(row) <= CI_STATUS:
            continue
        if row[CI_ORDER] is None or str(row[CI_ORDER]).strip() == "":
            continue
        oid = str(row[CI_ORDER]).strip()
        sc = _sc(row[CI_STATUS])
        if sc in exclude_statuses:
            continue
        if excl_test:
            blob = " ".join(str(row[c] or "") for c in (CI_SENDER, CI_RECEIVER)).lower()
            if any(k.lower() in blob for k in test_kw):
                continue

        include = False
        po = str(row[CI_PO] or "").strip().upper()
        # Check 1: currently at DVC/HUB
        if "DVC" in po or "HUB" in po:
            include = True
        # Check 2: passed through MEGA hub and still in transit
        if not include and len(row) > CI_ACTION_PO_306:
            action_po = str(row[CI_ACTION_PO_306] or "").strip().upper()
            if "MEGA" in action_po and sc in transit_statuses:
                include = True

        if not include:
            continue
        if oid in seen_orders:
            continue
        seen_orders.add(oid)
        filtered.append(row)'''

new_filter = '''    filtered = []
    for row in data_rows:
        if not row or len(row) <= CI_STATUS:
            continue
        if row[CI_ORDER] is None or str(row[CI_ORDER]).strip() == "":
            continue
        sc = _sc(row[CI_STATUS])
        if sc in exclude_statuses:
            continue
        if excl_test:
            blob = " ".join(str(row[c] or "") for c in (CI_SENDER, CI_RECEIVER)).lower()
            if any(k.lower() in blob for k in test_kw):
                continue
        po = str(row[CI_PO] or "").strip().upper()
        if not ("MEGA" in po or "HUB" in po or "DVC" in po):
            continue
        filtered.append(row)'''

content2 = content2.replace(old_filter, new_filter)

with open(path2, 'w', encoding='utf-8') as f:
    f.write(content2)
print("DONE mega_detail.py")
