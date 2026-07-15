import os

log_path = "c:/Users/DELL/Desktop/daily_push/bot.log"
if os.path.exists(log_path):
    print("Searching bot.log for today's push/send actions...")
    with open(log_path, "r", encoding="utf-8", errors="ignore") as f:
        for line in f:
            if "2026-07-13" in line:
                ll = line.lower()
                if any(x in ll for x in ("send", "push", "group", "success", "error", "fail")):
                    # skip Updater/Application startup messages to keep it clean
                    if "updater" not in ll and "application" not in ll and "getupdates" not in ll:
                        print(line.strip())
else:
    print("bot.log not found.")
