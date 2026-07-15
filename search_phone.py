import pandas as pd
import os

target_phones = ["060284059", "0887075591", "60284059", "887075591"]

files_to_check = [
    "test_detail.xlsx",
    "../test_detail.xlsx",
    "../fresh_detail.xlsx",
    "../../fresh_detail.xlsx"
]

print("Searching for phone numbers...")

found_any = False

for fpath in files_to_check:
    if not os.path.exists(fpath):
        continue
    print(f"\nChecking: {fpath}")
    try:
        # Load the excel file
        # We read it in chunks or fully
        df = pd.read_excel(fpath)
        
        # We will search every column for these numbers as string matches
        for col in df.columns:
            # Convert column to string
            col_str = df[col].astype(str)
            for phone in target_phones:
                matches = df[col_str.str.contains(phone, na=False, regex=False)]
                if not matches.empty:
                    found_any = True
                    print(f"  ★ Found match in column '{col}' for phone '{phone}':")
                    for idx, row in matches.iterrows():
                        order_id = row.get("ORDER ID") or row.get("Order ID") or row.get("order id")
                        sender = row.get("SENDER") or row.get("Sender")
                        receiver = row.get("RECEIVER") or row.get("Receiver")
                        status = row.get("CURRENT STATUS") or row.get("Current Status")
                        post_office = row.get("CURRENT POST OFFICE") or row.get("Current Post Office")
                        created_date = row.get("CREATED DATE") or row.get("Created Date")
                        print(f"    - Row {idx}: Order ID: {order_id} | Created: {created_date} | Status: {status} | PO: {post_office} | Sender: {sender} | Receiver: {receiver}")
    except Exception as e:
        print(f"  Error reading {fpath}: {e}")

if not found_any:
    print("\nNo matches found in any checked files.")
