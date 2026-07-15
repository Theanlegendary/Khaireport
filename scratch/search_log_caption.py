import os

log_path = "c:/Users/DELL/Desktop/daily_push/bot.log"
if os.path.exists(log_path):
    print("Searching bot.log for report counts from today...")
    with open(log_path, "r", encoding="utf-8", errors="ignore") as f:
        for line in f:
            if "2026-07-13" in line:
                line_lower = line.lower()
                if any(x in line_lower for x in ("pickup:", "delivery:", "pending:", "total:", "grand total:")):
                    print(line.strip())
else:
    print("bot.log not found.")
