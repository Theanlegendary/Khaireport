"""Update mega_detail.py to also include orders that passed through MEGA1/DVCMEGA1."""
import os, re

HERE = os.path.dirname(os.path.abspath(__file__))
path = os.path.join(HERE, 'mega_detail.py')

with open(path, 'r', encoding='utf-8') as f:
    content = f.read()

# Add COL_ACTION_PO_306
if 'CI_ACTION_PO_306' not in content:
    content = content.replace(
        'CI_STATUS = 23',
        'CI_STATUS = 23\nCI_ACTION_PO_306 = 35'
    )

# Replace the filtering logic
old_filter = '''    filtered = []
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

new_filter = '''    transit_statuses = {"210", "300", "302", "306", "309", "310", "311", "500"}
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

content = content.replace(old_filter, new_filter)

with open(path, 'w', encoding='utf-8') as f:
    f.write(content)
print("DONE")
