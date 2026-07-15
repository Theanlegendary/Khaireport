import os
import json

with open("c:/Users/DELL/Desktop/daily_push/reports_today.json") as f:
    data = json.load(f)

today_paths = data.get("2026-07-13", [])
print(f"Found {len(today_paths)} paths for today.")

for path in today_paths:
    if os.path.exists(path):
        files = os.listdir(path)
        mtime = os.path.getmtime(path)
        mtime_dt = datetime = datetime = datetime = None
        # let's just use os.path.getmtime and format it
        from datetime import datetime
        mtime_dt = datetime.fromtimestamp(mtime)
        print(f"Path: {path} | Modified: {mtime_dt} | Files: {files}")
    else:
        print(f"Path: {path} | DOES NOT EXIST")
