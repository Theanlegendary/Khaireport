import os

log_path = "c:/Users/DELL/Desktop/daily_push/bot.log"
if os.path.exists(log_path):
    print("Listing all log entries from 2026-07-13...")
    with open(log_path, "r", encoding="utf-8", errors="ignore") as f:
        today_lines = [line.strip() for line in f if "2026-07-13" in line]
    
    print(f"Total lines today: {len(today_lines)}")
    print("First 20 lines:")
    for line in today_lines[:20]:
        print("  ", line)
    print("Last 20 lines:")
    for line in today_lines[-20:]:
        print("  ", line)
else:
    print("bot.log not found.")
