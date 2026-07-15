import os
import re
import pandas as pd

# The raw target digits (without leading zero or country code)
targets = ["60284059", "887075591"]
downloads_dir = r"C:\Users\DELL\Downloads\Telegram Desktop"

print("Scanning Excel files with smart phone number matching...")
found_any = False

files = [f for f in os.listdir(downloads_dir) if f.endswith('.xlsx') and not f.startswith('~$')]

# Sort files by modified time (newest first)
files_with_time = []
for f in files:
    fpath = os.path.join(downloads_dir, f)
    try:
        files_with_time.append((fpath, os.path.getmtime(fpath)))
    except Exception:
        pass
files_with_time.sort(key=lambda x: x[1], reverse=True)

print(f"Total files: {len(files_with_time)}")

# We will scan only the 250 newest files first to keep it fast, then scan more if nothing found.
for idx, (fpath, mtime) in enumerate(files_with_time):
    fname = os.path.basename(fpath)
    fname_safe = fname.encode('ascii', errors='replace').decode('ascii')
    
    # We want to be thorough but fast
    if idx > 300 and found_any:
        # Stop if we found matches in the newest 300 files
        break
        
    try:
        xl = pd.ExcelFile(fpath)
        for sheet in xl.sheet_names:
            df = xl.parse(sheet)
            for col in df.columns:
                col_str = df[col].astype(str)
                for phone in targets:
                    # Clean the cell values to keep only digits
                    # e.g. "+855 060-284-059" -> "855060284059"
                    # We check if the target string (e.g. "60284059") is a substring of the cleaned digits
                    def is_match(val):
                        digits = re.sub(r'\D', '', str(val))
                        return phone in digits

                    mask = col_str.apply(is_match)
                    matches = df[mask]
                    
                    if not matches.empty:
                        found_any = True
                        print(f"\n★ SMART MATCH FOUND in file: {fname_safe} (Sheet: {sheet}, Column: {col}) for phone '{phone}':")
                        for idx_row, row in matches.iterrows():
                            details = {}
                            for k, v in row.items():
                                if pd.notna(v):
                                    k_safe = str(k).encode('ascii', errors='replace').decode('ascii')
                                    v_safe = str(v).encode('ascii', errors='replace').decode('ascii')
                                    details[k_safe] = v_safe
                            print(f"  - Row {idx_row}: {details}")
    except Exception as e:
        e_safe = str(e).encode('ascii', errors='replace').decode('ascii')
        # print(f"  Error reading {fname_safe}: {e_safe}")

print("\nScan complete.")
