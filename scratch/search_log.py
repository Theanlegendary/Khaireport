import os

log_path = "c:/Users/DELL/Desktop/daily_push/bot.log"
if os.path.exists(log_path):
    print("Searching bot.log for entries from today...")
    with open(log_path, "r", encoding="utf-8", errors="ignore") as f:
        lines = f.readlines()
    
    # We want to print lines containing today's date "2026-07-13" or "13/07/2026"
    count = 0
    for line in lines:
        if "2026-07-13" in line or "13/07/2026" in line or "Overall counts" in line:
            print(line.strip())
            count += 1
            if count > 200:
                print("... truncated ...")
                break
else:
    print("bot.log not found.")
