import os
import pandas as pd

target_subs = ["60284059", "887075591", "284059", "7075591"]
base_dir = r"C:\Users\DELL\Downloads\Telegram Desktop\DataPusher 2\DataPusher"

print(f"Scanning ALL files (CSV, TXT, JSON, etc.) in: {base_dir}")

found_any = False

for root, dirs, files in os.walk(base_dir):
    for f in files:
        fpath = os.path.join(root, f)
        # Skip large excel/zip files unless they are text
        if f.endswith(('.zip', '.png', '.jpg', '.jpeg', '.pyc', '.exe')):
            continue
            
        # If excel, we already scanned it, but we can do a raw text scan on CSV/TXT/JSON
        if f.endswith(('.csv', '.txt', '.json', '.html', '.env')):
            try:
                with open(fpath, 'r', encoding='utf-8', errors='ignore') as file_in:
                    lines = file_in.readlines()
                    for idx, line in enumerate(lines):
                        for sub in target_subs:
                            if sub in line:
                                found_any = True
                                print(f"\n★ FOUND MATCH in text file: {fpath} (Line {idx+1}) for '{sub}':")
                                print(f"  {line.strip()[:150]}")
            except Exception as e:
                print(f"  Error reading {fpath}: {e}")

if not found_any:
    print("\nNo matches found in any text files.")
