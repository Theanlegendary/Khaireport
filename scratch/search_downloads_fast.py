import os
from datetime import datetime, date

today_dt = date(2026, 7, 13)
base_dir = r"C:\Users\DELL\Downloads"

print(f"Scanning {base_dir} for Excel files modified today...")
found_files = []

for entry in os.scandir(base_dir):
    if entry.is_file() and entry.name.endswith('.xlsx') and not entry.name.startswith('~$'):
        try:
            mtime = entry.stat().st_mtime
            mdate = datetime.fromtimestamp(mtime).date()
            if mdate == today_dt:
                mtime_dt = datetime.fromtimestamp(mtime)
                found_files.append((entry.path, mtime_dt))
                print(f"File: {entry.path} | Modified: {mtime_dt}")
        except Exception:
            pass

# Also scan subdirectories under C:\Users\DELL\Downloads but skip "Telegram Desktop" to avoid duplication/heavy folders
for entry in os.scandir(base_dir):
    if entry.is_dir() and entry.name.lower() not in ("telegram desktop", "datapusher 2", "datapusher"):
        # Scan 1 level deep
        try:
            for sub_entry in os.scandir(entry.path):
                if sub_entry.is_file() and sub_entry.name.endswith('.xlsx') and not sub_entry.name.startswith('~$'):
                    mtime = sub_entry.stat().st_mtime
                    mdate = datetime.fromtimestamp(mtime).date()
                    if mdate == today_dt:
                        mtime_dt = datetime.fromtimestamp(mtime)
                        found_files.append((sub_entry.path, mtime_dt))
                        print(f"File: {sub_entry.path} | Modified: {mtime_dt}")
        except Exception:
            pass

print(f"Total files found: {len(found_files)}")
