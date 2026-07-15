import os
from datetime import datetime, date

today_dt = date(2026, 7, 13)
temp_dir = r"C:\Users\DELL\AppData\Local\Temp"

print("Scanning Temp directory for today's Excel files...")
found_files = []

for root, dirs, files in os.walk(temp_dir):
    # skip deep folders to be fast
    if len(root.split(os.sep)) > 8:
        continue
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
