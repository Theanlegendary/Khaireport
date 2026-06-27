import sys
import pandas as pd
import json
import os

sys.stdout.reconfigure(encoding='utf-8')

path = "Bill test 17.06.xlsx"
if os.path.exists(path):
    df = pd.read_excel(path, dtype=str)
    print("Total rows in Excel:", len(df))
    
    order_col = next(
        (
            c for c in df.columns
            if "order" in str(c).lower()
            or "phi" in str(c).lower()
            or "shipment" in str(c).lower()
        ),
        df.columns[0] if len(df.columns) else None,
    )
    print("Detected column:", order_col)
    if order_col:
        all_vals = df[order_col].tolist()
        non_empty = [v for v in all_vals if str(v).strip() and str(v).strip().lower() != "nan"]
        print("Total non-empty values in column:", len(non_empty))
        
        unique_vals = set(non_empty)
        print("Total UNIQUE IDs:", len(unique_vals))
        
        # Let's count duplicates
        duplicates = len(non_empty) - len(unique_vals)
        print("Duplicate values in column:", duplicates)
