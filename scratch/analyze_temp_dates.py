import os
from datetime import datetime, date

temp_dir = r"C:\Users\DELL\AppData\Local\Temp"
today_dt = date(2026, 7, 13)

folders = []
for f in os.listdir(temp_dir):
    fpath = os.path.join(temp_dir, f)
    if os.path.isdir(fpath) and (f.startswith("total_") or f.startswith("push_") or "report" in f.lower()):
        try:
            mtime = os.path.getmtime(fpath)
            mtime_dt = datetime.fromtimestamp(mtime)
            if mtime_dt.date() == today_dt:
                folders.append((fpath, mtime_dt))
        except Exception:
            pass

# Sort by modification time ascending
folders.sort(key=lambda x: x[1])
for fpath, mtime in folders:
    print(f"Folder: {fpath} | Modified: {mtime}")
