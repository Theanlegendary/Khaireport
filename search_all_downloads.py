import os
import re
import pandas as pd

targets = ["60284059", "887075591"]
downloads_dir = r"C:\Users\DELL\Downloads\Telegram Desktop"

print("Scanning all files (XLSX, XLS, XLSB, CSV, TXT, JSON) in downloads...")
found_any = False

files = os.listdir(downloads_dir)
files_with_time = []
for f in files:
    fpath = os.path.join(downloads_dir, f)
    if os.path.isdir(fpath) or f.startswith('~$'):
        continue
    try:
        files_with_time.append((fpath, os.path.getmtime(fpath)))
    except Exception:
        pass

# Sort files by modified time (newest first)
files_with_time.sort(key=lambda x: x[1], reverse=True)
print(f"Total files to check: {len(files_with_time)}")

def check_df(df, fname_safe, source_desc):
    global found_any
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
                print(f"\n★ MATCH FOUND in {fname_safe} ({source_desc}, Column: {col}) for phone '{phone}':")
                for idx_row, row in matches.iterrows():
                    details = {}
                    for k, v in row.items():
                        if pd.notna(v):
                            k_safe = str(k).encode('ascii', errors='replace').decode('ascii')
                            v_safe = str(v).encode('ascii', errors='replace').decode('ascii')
                            details[k_safe] = v_safe
                    print(f"  - Row {idx_row}: {details}")

for idx, (fpath, mtime) in enumerate(files_with_time):
    fname = os.path.basename(fpath)
    fname_safe = fname.encode('ascii', errors='replace').decode('ascii')
    ext = os.path.splitext(fname)[1].lower()
    
    # We want to check all files, but print progress every 50 files
    if idx % 50 == 0:
        print(f"Progress: checked {idx}/{len(files_with_time)} files...")

    try:
        if ext == '.xlsx':
            xl = pd.ExcelFile(fpath)
            for sheet in xl.sheet_names:
                df = xl.parse(sheet)
                check_df(df, fname_safe, f"Sheet: {sheet}")
        elif ext == '.xls':
            xl = pd.ExcelFile(fpath, engine='xlrd')
            for sheet in xl.sheet_names:
                df = xl.parse(sheet)
                check_df(df, fname_safe, f"Sheet: {sheet}")
        elif ext == '.xlsb':
            xl = pd.ExcelFile(fpath, engine='pyxlsb')
            for sheet in xl.sheet_names:
                df = xl.parse(sheet)
                check_df(df, fname_safe, f"Sheet: {sheet}")
        elif ext == '.csv':
            # Try loading as CSV with different encodings
            for encoding in ['utf-8', 'latin-1', 'utf-16']:
                try:
                    df = pd.read_csv(fpath, encoding=encoding, low_memory=False)
                    check_df(df, fname_safe, "CSV File")
                    break
                except Exception:
                    continue
        elif ext in ['.txt', '.json', '.html', '.xml']:
            with open(fpath, 'r', encoding='utf-8', errors='ignore') as file_in:
                lines = file_in.readlines()
                for line_idx, line in enumerate(lines):
                    line_digits = re.sub(r'\D', '', line)
                    for phone in targets:
                        if phone in line_digits:
                            found_any = True
                            print(f"\n★ MATCH FOUND in text file: {fname_safe} (Line {line_idx+1}) for '{phone}':")
                            print(f"  {line.strip()[:150]}")
    except Exception as e:
        # Ignore errors to keep running
        pass

print(f"\nScan complete. Checked {len(files_with_time)} files.")
