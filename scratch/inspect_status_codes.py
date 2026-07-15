import pandas as pd
df = pd.read_excel("c:/Users/DELL/Desktop/daily_push/test_detail.xlsx")
# clean/extract status codes
if 'CURRENT STATUS' in df.columns:
    df['STATUS_CODE'] = df['CURRENT STATUS'].astype(str).str.strip().str.extract(r'^(\d{3})')[0]
    print("All status codes in dataset:")
    print(df['STATUS_CODE'].value_counts())
    
    # Check BANP001 status codes
    # wait, what column is the branch postcode? Let's check columns first
    print("Columns:", list(df.columns))
