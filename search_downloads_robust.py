import os
import pandas as pd

target_phones = ["060284059", "0887075591", "60284059", "887075591"]
downloads_dir = r"C:\Users\DELL\Downloads\Telegram Desktop"

print("Scanning all Excel files in downloads safely...")

found_any = False

files = [f for f in os.listdir(downloads_dir) if f.endswith('.xlsx') and not f.startswith('~$')]

# Sort files by modified time, newest first
files_with_time = []
for f in files:
    fpath = os.path.join(downloads_dir, f)
    try:
        files_with_time.append((fpath, os.path.getmtime(fpath)))
    except Exception:
        pass
files_with_time.sort(key=lambda x: x[1], reverse=True)

print(f"Total files found: {len(files_with_time)}")

for fpath, mtime in files_with_time:
    fname = os.path.basename(fpath)
    # Safe printing of filename
    fname_safe = fname.encode('ascii', errors='replace').decode('ascii')
    
    try:
        xl = pd.ExcelFile(fpath)
        for sheet in xl.sheet_names:
            df = xl.parse(sheet)
            for col in df.columns:
                col_str = df[col].astype(str)
                for phone in target_phones:
                    matches = df[col_str.str.contains(phone, na=False, regex=False)]
                    if not matches.empty:
                        found_any = True
                        print(f"\n★ MATCH FOUND in file: {fname_safe} (Sheet: {sheet}, Column: {col}) for phone '{phone}':")
                        for idx, row in matches.iterrows():
                            # Extract common fields safely
                            details = {}
                            for k, v in row.items():
                                if pd.notna(v):
                                    k_safe = str(k).encode('ascii', errors='replace').decode('ascii')
                                    v_safe = str(v).encode('ascii', errors='replace').decode('ascii')
                                    details[k_safe] = v_safe
                            print(f"  - Row {idx}: {details}")
    except Exception as e:
        # print error safely
        e_safe = str(e).encode('ascii', errors='replace').decode('ascii')
        print(f"  Error reading {fname_safe}: {e_safe}")

if not found_any:
    print("\nScan complete. No matches found.")
