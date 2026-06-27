import urllib.request
import json
import os
import sys
import pandas as pd
import tempfile
from datetime import datetime

# Add root folder to sys.path so we can import downloader
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.stdout.reconfigure(encoding='utf-8')

def main():
    print("Loading config...")
    with open("config.json", "r", encoding="utf-8") as f:
        cfg = json.load(f)
        
    tmpdir = tempfile.mkdtemp(prefix="puttra_")
    out_path = os.path.join(tmpdir, "latest_detail.xlsx")
    
    print("Downloading the latest details from API...")
    import downloader
    downloader.download_detail(cfg["api"], out_path, force_refresh=True)
    
    print("Loading Excel file...")
    xl = pd.ExcelFile(out_path)
    sheet = xl.sheet_names[0]
    # Check if there are other sheets containing order data
    for sname in xl.sheet_names:
        try:
            s = xl.parse(sname, nrows=3)
            if 'ORDER ID' in s.columns:
                sheet = sname
                break
        except Exception:
            continue
            
    df = xl.parse(sheet)
    print(f"Loaded sheet '{sheet}' with {len(df)} rows.")
    
    # Let's search for Ngorn Puttra (ne300012 or 855978843545)
    postman_pattern = "Ngorn Puttra|855978843545|ne300012"
    
    # We want to check all rows where he is the action user
    match_user = pd.Series(False, index=df.index)
    for col in df.columns:
        if 'user' in col.lower() or 'shipper' in col.lower() or 'action' in col.lower() or 'postman' in col.lower():
            match_user = match_user | df[col].astype(str).str.contains(postman_pattern, na=False, case=False)
            
    matched_df = df[match_user].copy()
    print(f"Found {len(matched_df)} total historical matching scans for Ngorn Puttra.")
    
    # Now, filter to scans on 26/06/2026 or 2026-06-26
    date_pattern = "26/06/2026|2026-06-26"
    match_date = pd.Series(False, index=matched_df.index)
    for col in matched_df.columns:
        if 'time' in col.lower() or 'date' in col.lower():
            match_date = match_date | matched_df[col].astype(str).str.contains(date_pattern, na=False)
            
    scans_26 = matched_df[match_date].copy()
    print(f"Found {len(scans_26)} scans on 26/06/2026.")
    
    # Let's display the unique order IDs and details
    if not scans_26.empty:
        # Display the table of columns
        display_cols = ['ORDER ID', 'CURRENT STATUS', 'CURRENT TIME', 'ACTION USER']
        cols_to_print = [c for c in display_cols if c in scans_26.columns]
        
        print("\n=== Order Scans on 26/06/2026 ===")
        print(scans_26[cols_to_print].to_string(index=False))
        
        # Also print detailed history for each of these unique order IDs to check if they have S402/S410 scans
        unique_ids = scans_26['ORDER ID'].unique().tolist()
        print(f"\nChecking tracking history for {len(unique_ids)} unique orders:")
        
        headers = {
            'Authorization': 'Bearer ' + cfg['api']['bearer_token'],
            'Accept': 'application/json, text/plain, */*',
            'User-Agent': 'Mozilla/5.0'
        }
        
        for oid in unique_ids:
            oid_str = str(oid)
            print(f"\nOrder ID: {oid_str}")
            r_track = requests.get('https://gw-express.metfone.com.kh/tms-tracking/api/v1/order-tracking', params={'order_id': oid_str}, headers=headers)
            if r_track.status_code == 200:
                trips = r_track.json().get('trackingTrips', [])
                for t in trips:
                    # check if Ngorn Puttra scanned this milestone
                    op_name = t.get('updatedBy', {}).get('name', '')
                    status_code = t.get('status', '').lstrip('S')
                    status_name = t.get('statusName', '')
                    updated_at = t.get('updatedAt', '')
                    desc = t.get('desc', '')
                    
                    is_puttra = any(w in op_name.lower() for w in ['puttra', 'ne300012'])
                    marker = "⭐️ [Puttra]" if is_puttra else "  [Other]"
                    print(f"  {marker} {updated_at[:19]} | S{status_code} ({status_name}) | User: {op_name} | Desc: {desc}")
            else:
                print(f"  Failed to get tracking: {r_track.status_code}")
    else:
        print("No scans found on 26/06/2026.")

if __name__ == "__main__":
    import requests
    main()
