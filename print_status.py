import pandas as pd
df = pd.read_excel('test_detail.xlsx')

with open('status_output.txt', 'w', encoding='utf-8') as f_out:
    f_out.write("Unique values of 'CURRENT STATUS':\n")
    for val in df['CURRENT STATUS'].unique():
        f_out.write(f"{repr(val)}\n")

    f_out.write("\nAre there columns with user/messenger details?\n")
    for col in df.columns:
        if 'user' in col.lower() or 'action' in col.lower() or 'note' in col.lower() or 'status' in col.lower():
            f_out.write(f"{col}\n")
            vals = df[col].dropna().unique()[:5]
            f_out.write(f"  Values: {[repr(v) for v in vals]}\n")
print("Done writing to status_output.txt")
