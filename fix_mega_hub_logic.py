"""Update build_mega_pivot to also count orders that passed through MEGA1/DVCMEGA1
but are still in transit (stuck after leaving mega hub)."""
import os, re

HERE = os.path.dirname(os.path.abspath(__file__))
path = os.path.join(HERE, 'pivot.py')

with open(path, 'r', encoding='utf-8') as f:
    content = f.read()

# Add column index for ACTION POST OFFICE (col 35)
if 'COL_ACTION_PO_306' not in content:
    # Add after COL_RECEIVER line
    content = content.replace(
        'COL_RECEIVER = 4',
        'COL_RECEIVER = 4\nCOL_ACTION_PO_306 = 35  # ACTION POST OFFICE for status 306'
    )

# Replace build_mega_pivot to also check col 35 for MEGA1/DVCMEGA1
old_start = content.find('def build_mega_pivot(rows, pivot_cfg, zone_cfg):')
old_end = content.find('\n\ndef export_mega_pivot')

old_func = content[old_start:old_end]

new_func = '''def build_mega_pivot(rows, pivot_cfg, zone_cfg):
    """
    Builds a pivot tree for MEGA check:
      tree: dict[hub_code][(month, day)] = count
      day_keys: sorted list of (month, day) tuples present in the data
      urgent_tree: dict[hub_code] = urgent_count (orders >1 day old)
    
    Logic:
      - DVCZ1-5: orders where CURRENT POST OFFICE (col 15) contains DVC/HUB
      - MEGA1/DVCMEGA1: orders where ACTION POST OFFICE 306 (col 35) contains MEGA
        AND status is still transit (not delivered/done)
    """
    exclude_test = pivot_cfg.get("exclude_test", False)
    test_keywords = pivot_cfg.get("test_keywords", ["test"])
    
    exclude_statuses = {"410", "201", "520", "99", "100", "-99"}
    # Statuses that mean parcel is still in transit (not yet at receiver store)
    transit_statuses = {"210", "300", "302", "306", "309", "310", "311", "500"}

    tree = defaultdict(lambda: defaultdict(int))
    urgent_tree = defaultdict(int)
    day_keys_seen = set()
    seen_orders = set()  # avoid double-counting
    today = datetime.now().date()

    for row in rows:
        if not row or row[COL_ORDER_ID] in (None, ""):
            continue
        order_id = str(row[COL_ORDER_ID]).strip()
        status_code = _status_code(row[COL_CURRENT_STATUS])
        if status_code in exclude_statuses:
            continue
        if exclude_test and _is_test_row(row, test_keywords):
            continue

        po = str(row[COL_CURRENT_PO] or "").strip()
        po_upper = po.upper()
        
        hub_label = None

        # Check 1: CURRENT POST OFFICE contains DVC or HUB (currently stuck at hub)
        if "DVC" in po_upper or "HUB" in po_upper:
            hub_label = po
        
        # Check 2: ACTION POST OFFICE 306 (col 35) contains MEGA
        # AND status is still transit (not yet delivered to receiver store)
        if hub_label is None and len(row) > COL_ACTION_PO_306:
            action_po = str(row[COL_ACTION_PO_306] or "").strip().upper()
            if "MEGA" in action_po and status_code in transit_statuses:
                hub_label = str(row[COL_ACTION_PO_306] or "").strip()

        if hub_label is None:
            continue
        if order_id in seen_orders:
            continue
        seen_orders.add(order_id)

        month, day = _parse_day(row[COL_CREATED_DATE])
        if day is None:
            continue

        key = (month, day)
        tree[hub_label][key] += 1
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
            urgent_tree[hub_label] += 1

    day_keys = sorted(day_keys_seen)
    return tree, day_keys, urgent_tree'''

content = content[:old_start] + new_func + content[old_end:]

with open(path, 'w', encoding='utf-8') as f:
    f.write(content)
print("DONE")
