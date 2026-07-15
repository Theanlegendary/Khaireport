import openpyxl

fpath = r"c:\Users\DELL\Desktop\daily_push\test_out\Report_All_13_07_2026.xlsx"
wb = openpyxl.load_workbook(fpath)

with open("c:/Users/DELL/Desktop/daily_push/scratch/afternoon_debug_output.txt", "w", encoding="utf-8") as out:
    for sname in wb.sheetnames:
        out.write(f"\n=========================================\n")
        out.write(f"Sheet Name: {sname}\n")
        ws = wb[sname]
        
        # Print first 10 rows
        for r in range(1, min(20, ws.max_row + 1)):
            row_vals = [str(ws.cell(r, c).value) if ws.cell(r, c).value is not None else "" for c in range(1, ws.max_column + 1)]
            out.write(f"Row {r:02d}: {', '.join(row_vals)}\n")
            
print("SUCCESS - Written to scratch/afternoon_debug_output.txt")
