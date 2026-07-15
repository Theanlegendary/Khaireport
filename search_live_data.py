import os
import json
import re
import time
import pandas as pd
import downloader

# The raw target digits
targets = ["60284059", "887075591"]

print("Loading config...")
with open("config.json", encoding="utf-8") as f:
    cfg = json.load(f)

src_path = "live_export_find.xlsx"

print("Downloading live detail file from API with retries...")
download_ok = False
for attempt in range(5):
    try:
        print(f"Attempt {attempt+1}/5...")
        downloader.download_detail(cfg["api"], src_path, force_refresh=True)
        print("Download completed successfully!")
        download_ok = True
        break
    except Exception as e:
        print(f"Error on attempt {attempt+1}: {e}")
        time.sleep(3)

if not download_ok:
    print("Failed to download live data after 5 attempts.")
    if os.path.exists(src_path):
        print("Using existing local live_export_find.xlsx file.")
    else:
        exit(1)

print(f"Scanning downloaded file: {src_path}")
try:
    xl = pd.ExcelFile(src_path)
    sheet = xl.sheet_names[0]
    if len(xl.sheet_names) > 1:
        for sname in xl.sheet_names:
            try:
                s = xl.parse(sname, nrows=3)
                if 'ORDER ID' in s.columns:
                    sheet = sname
                    break
            except Exception:
                continue
    
    df = xl.parse(sheet)
    print(f"Loaded sheet: {sheet} with {len(df)} rows.")
    
    found_any = False
    for col in df.columns:
        col_str = df[col].astype(str)
        for phone in targets:
            def is_match(val):
                digits = re.sub(r'\D', '', str(val))
                return phone in digits

            mask = col_str.apply(is_match)
            matches = df[mask]
            
            if not matches.empty:
                found_any = True
                print(f"\n★ MATCH FOUND in column '{col}' for phone '{phone}':")
                for idx_row, row in matches.iterrows():
                    details = {}
                    for k, v in row.items():
                        if pd.notna(v):
                            details[str(k)] = str(v)
                    print(f"  - Row {idx_row}:")
                    for k_d, v_d in details.items():
                        print(f"    {k_d}: {v_d}")
                        
    if not found_any:
        print("\nNo matching orders found in the live data.")
except Exception as e:
    print(f"Error reading file: {e}")
