import openpyxl

fpath = r"c:\Users\DELL\Desktop\daily_push\test_out\Report_All_13_07_2026.xlsx"
wb = openpyxl.load_workbook(fpath)

for sname in wb.sheetnames:
    # safe print sheet name by encoding it to utf-8 and printing repr
    print(f"Sheet Name Repr: {repr(sname)}")
    ws = wb[sname]
    for r in range(1, 6):
        row_vals = [ws.cell(r, c).value for c in range(1, 10)]
        # safe print row values
        safe_vals = [str(x).encode('utf-8', errors='replace').decode('utf-8') if x is not None else None for x in row_vals]
        print(f"  Row {r}: {safe_vals}")
