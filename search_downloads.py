import os
import pandas as pd

target_phones = ["060284059", "0887075591", "60284059", "887075591"]
downloads_dir = r"C:\Users\DELL\Downloads\Telegram Desktop"

print(f"Scanning .xlsx files directly in: {downloads_dir}")

found_any = False

# List files in the downloads directory
files = [f for f in os.listdir(downloads_dir) if f.endswith('.xlsx') and not f.startswith('~$')]

# Sort files by modified time, newest first
files_with_time = []
for f in files:
    fpath = os.path.join(downloads_dir, f)
    files_with_time.append((fpath, os.path.getmtime(fpath)))
files_with_time.sort(key=lambda x: x[1], reverse=True)

print(f"Found {len(files_with_time)} Excel files. Scanning starting from the newest...")

for fpath, mtime in files_with_time:
    fname = os.path.basename(fpath)
    print(f"Checking: {fname}")
    try:
        xl = pd.ExcelFile(fpath)
        for sheet in xl.sheet_names:
            df = xl.parse(sheet)
            for col in df.columns:
                # convert to string and search
                col_str = df[col].astype(str)
                for phone in target_phones:
                    # check for substring match
                    matches = df[col_str.str.contains(phone, na=False, regex=False)]
                    if not matches.empty:
                        found_any = True
                        print(f"\n★ FOUND MATCH in file: {fname} (Sheet: {sheet}, Column: {col}) for phone '{phone}':")
                        for idx, row in matches.iterrows():
                            order_id = row.get("ORDER ID") or row.get("Order ID") or row.get("order id") or row.get("Mã vận đơn")
                            sender = row.get("SENDER") or row.get("Sender") or row.get("Người gửi") or row.get("SENDER PHONE") or row.get("RECEIVER PHONE")
                            receiver = row.get("RECEIVER") or row.get("Receiver") or row.get("Người nhận")
                            status = row.get("CURRENT STATUS") or row.get("Current Status") or row.get("Trạng thái hiện tại")
                            po = row.get("CURRENT POST OFFICE") or row.get("Current Post Office") or row.get("Bưu cục hiện tại")
                            created_date = row.get("CREATED DATE") or row.get("Created Date") or row.get("Ngày tạo")
                            
                            # Print all values of the row to be extremely helpful
                            row_details = {k: v for k, v in row.items() if pd.notna(v)}
                            print(f"  - Row {idx}: Order ID: {order_id} | Created: {created_date} | Status: {status} | PO: {po}")
                            print(f"    Full Row Data: {row_details}")
    except Exception as e:
        print(f"  Error reading {fname}: {e}")

if not found_any:
    print("\nNo matches found in any Excel file in Telegram Desktop downloads.")
