import os
from datetime import datetime, date

today_dt = date(2026, 7, 13)
base_dirs = [
    r"c:\Users\DELL\Desktop",
    r"c:\Users\DELL\Documents",
    r"C:\Users\DELL\Downloads",
]

print("Scanning for all Excel files modified today on Desktop, Documents, and Downloads...")
found_files = []

for base_dir in base_dirs:
    if not os.path.exists(base_dir):
        continue
    for root, dirs, files in os.walk(base_dir):
        # Limit recursion depth to keep it fast
        depth = len(root.split(os.sep)) - len(base_dir.split(os.sep))
        if depth > 3:
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
