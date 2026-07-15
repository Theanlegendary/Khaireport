import pandas as pd
df = pd.read_excel("c:/Users/DELL/Desktop/daily_push/test_detail.xlsx")
ref_df = pd.read_csv("c:/Users/DELL/Desktop/daily_push/post_office_lookup.csv")

# Clean keys
df['STATUS_CODE'] = df['CURRENT STATUS'].astype(str).str.strip().str.extract(r'^(\d{3})')[0]
df['_po_key'] = df['CURRENT POST OFFICE'].astype(str).str.strip().str.upper()
ref_df['_ref_key'] = ref_df['current_post_office'].astype(str).str.strip().str.upper()

# Merge
dm = pd.merge(df, ref_df, left_on='_po_key', right_on='_ref_key', how='left')
dm['POST OFFICE_HANDLE'] = dm['post_office_handle'].fillna(dm['CURRENT POST OFFICE']).astype(str).str.strip().str.upper()

# Find handles with 420 or 472
green_handles = dm[dm['STATUS_CODE'].isin(['420', '472'])]['POST OFFICE_HANDLE'].unique()
print("Handles with 420 or 472 (Green):", list(green_handles))

# Find handles with 430, 460, 500
orange_handles = dm[dm['STATUS_CODE'].isin(['430', '460', '500'])]['POST OFFICE_HANDLE'].unique()
print("Handles with 430, 460, 500 (Orange):", list(orange_handles))
