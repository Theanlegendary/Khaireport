import os
from datetime import datetime, date

today_dt = date(2026, 7, 13)
base_dir = r"C:\Users\DELL\Downloads\Telegram Desktop\DataPusher 2\DataPusher"

print(f"Scanning {base_dir} recursively for Excel files modified today...")
found_files = []

for root, dirs, files in os.walk(base_dir):
    for f in files:
        if f.endswith('.xlsx') and not f.startswith('~$'):
            fpath = os.path.join(root, f)
            try:
                mtime = os.path.getmtime(fpath)
                mdate = datetime.fromtimestamp(mtime).date()
                if mdate == today_dt:
                    mtime_dt = datetime.fromtimestamp(mtime)
                    found_files.append((fpath, mtime_dt))
                    print(f"File: {fpath} | Modified: {mtime_dt}")
            except Exception:
                pass

print(f"Total files found: {len(found_files)}")
