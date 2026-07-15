import os
import pandas as pd

target_phones = ["060284059", "0887075591", "60284059", "887075591"]
base_dir = r"C:\Users\DELL\Downloads\Telegram Desktop\DataPusher 2\DataPusher"

print(f"Scanning directory: {base_dir}")

found_any = False

for root, dirs, files in os.walk(base_dir):
    for f in files:
        if f.endswith('.xlsx') and not f.startswith('~$'):
            fpath = os.path.join(root, f)
            print(f"Checking: {fpath}")
            try:
                # Load Excel
                xl = pd.ExcelFile(fpath)
                for sheet in xl.sheet_names:
                    df = xl.parse(sheet)
                    for col in df.columns:
                        col_str = df[col].astype(str)
                        for phone in target_phones:
                            matches = df[col_str.str.contains(phone, na=False, regex=False)]
                            if not matches.empty:
                                found_any = True
                                print(f"\n★ FOUND MATCH in file: {fpath} (Sheet: {sheet}, Column: {col}) for phone '{phone}':")
                                for idx, row in matches.iterrows():
                                    order_id = row.get("ORDER ID") or row.get("Order ID") or row.get("order id") or row.get("Mã vận đơn")
                                    sender = row.get("SENDER") or row.get("Sender") or row.get("Người gửi")
                                    receiver = row.get("RECEIVER") or row.get("Receiver") or row.get("Người nhận")
                                    status = row.get("CURRENT STATUS") or row.get("Current Status") or row.get("Trạng thái hiện tại")
                                    po = row.get("CURRENT POST OFFICE") or row.get("Current Post Office") or row.get("Bưu cục hiện tại")
                                    created_date = row.get("CREATED DATE") or row.get("Created Date") or row.get("Ngày tạo")
                                    print(f"  - Row {idx}: Order ID: {order_id} | Created: {created_date} | Status: {status} | PO: {po} | Sender: {sender} | Receiver: {receiver}")
            except Exception as e:
                print(f"  Error reading {fpath}: {e}")

if not found_any:
    print("\nNo matches found in any Excel file.")
