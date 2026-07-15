import pandas as pd

# Load data
df = pd.read_excel("c:/Users/DELL/Desktop/daily_push/test_detail.xlsx")
ref_df = pd.read_csv("c:/Users/DELL/Desktop/daily_push/post_office_lookup.csv")

# Clean keys
df['STATUS_CODE'] = df['CURRENT STATUS'].astype(str).str.strip().str.extract(r'^(\d{3})')[0]
df['_po_key'] = df['CURRENT POST OFFICE'].astype(str).str.strip().str.upper()
ref_df['_ref_key'] = ref_df['current_post_office'].astype(str).str.strip().str.upper()

# Merge
dm = pd.merge(df, ref_df, left_on='_po_key', right_on='_ref_key', how='left')
dm['POST OFFICE_HANDLE'] = dm['post_office_handle'].fillna(dm['CURRENT POST OFFICE']).astype(str).str.strip().str.upper()

# Clean STATUS_CODE drop (completed ones)
dm = dm[~dm['STATUS_CODE'].isin(['410', '201', '520'])].copy()

# Filter for PNPP002
dpn = dm[dm['POST OFFICE_HANDLE'] == 'PNPP002']
print(f"Total rows for PNPP002 in dataset: {len(dpn)}")

# Map statuses
import json
with open("c:/Users/DELL/Desktop/daily_push/config.json") as f:
    cfg = json.load(f)
status_map = {}
for r in cfg.get("reports", []):
    label = 'Pickup' if 'pickup' in r.get('key', '').lower() else \
            'Delivery' if 'delivery' in r.get('key', '').lower() else \
            'Pending' if 'pending' in r.get('key', '').lower() else r.get('is_label', '')
    for sc in r.get("status_codes", []):
        status_map[str(sc).strip()] = label

dpn['_report_class'] = dpn['STATUS_CODE'].map(status_map).fillna("Unknown")
print(dpn['_report_class'].value_counts())
