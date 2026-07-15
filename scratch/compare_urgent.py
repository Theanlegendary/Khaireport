import openpyxl

file_morning = r"c:\Users\DELL\Desktop\daily_push - Copy\test_out\Report_All_13_07_2026.xlsx"
file_afternoon = r"C:\Users\DELL\Downloads\Telegram Desktop\Total_13.07_14H32.xlsx"

# Helper to find date columns and their month/day mapping
def get_date_columns_map(ws):
    # Row 2 contains months (e.g. 'June', 'July'), Row 3 contains days
    date_cols = {} # {col_idx: (month, day)}
    
    current_month = None
    for c in range(1, ws.max_column + 1):
        month_val = ws.cell(2, c).value
        if month_val is not None:
            val_str = str(month_val).strip()
            if val_str in ("June", "July"):
                current_month = val_str
                
        day_val = ws.cell(3, c).value
        if day_val is not None and str(day_val).strip().isdigit() and current_month:
            date_cols[c] = (current_month, int(str(day_val).strip()))
            
    return date_cols

def is_urgent(month, day):
    # Overdue on July 13th means created on or before July 11th
    if month == "June":
        return True # Any June order is overdue on July 13th
    elif month == "July" and day <= 11:
        return True
    return False

# 1. Parse Morning (3 sheets)
def get_morning_urgent():
    wb = openpyxl.load_workbook(file_morning)
    counts = {"Pickup": 0, "Delivery": 0, "Pending": 0}
    for sname in wb.sheetnames:
        ws = wb[sname]
        key = "Pending"
        if "pickup" in sname.lower():
            key = "Pickup"
        elif "delivery" in sname.lower() or "deliver" in sname.lower():
            key = "Delivery"
            
        date_cols = get_date_columns_map(ws)
        
        # Find order_id column header
        order_col_idx = 4
        for c in range(1, ws.max_column + 1):
            cell_val = ws.cell(2, c).value
            if cell_val and any(h in str(cell_val).lower() for h in ('order id', 'លេខបុង')):
                order_col_idx = c
                break
        
        # Count urgent
        urgent_cnt = 0
        for r in range(4, ws.max_row + 1):
            order_id = ws.cell(r, order_col_idx).value
            if order_id is not None and str(order_id).strip().isdigit():
                created_date = None
                for c_idx, (m, d) in date_cols.items():
                    cell_val = ws.cell(r, c_idx).value
                    if cell_val == 1 or str(cell_val).strip() == '1':
                        created_date = (m, d)
                        break
                if created_date and is_urgent(created_date[0], created_date[1]):
                    urgent_cnt += 1
        counts[key] = urgent_cnt
    return counts

# 2. Parse Afternoon (Stacked Total Sheet)
def get_afternoon_urgent():
    wb = openpyxl.load_workbook(file_afternoon)
    ws = wb.active
    counts = {"Pickup": 0, "Delivery": 0, "Pending": 0}
    
    current_section = None
    date_cols = {}
    
    r = 1
    while r <= ws.max_row:
        val_a = ws.cell(r, 1).value
        val_a_str = str(val_a).strip() if val_a is not None else ""
        
        # Check for section start
        if "PICKUP" in val_a_str.upper() or "អីវ៉ាន់ត្រូវយក" in val_a_str:
            current_section = "Pickup"
            date_cols = {}
            r += 1
            continue
        elif "DELIVERY" in val_a_str.upper() or "អីវ៉ាន់ត្រូវដឹក" in val_a_str:
            current_section = "Delivery"
            date_cols = {}
            r += 1
            continue
        elif "PENDING" in val_a_str.upper() or "អីវ៉ាន់កំពុងរង់ចាំ" in val_a_str:
            current_section = "Pending"
            date_cols = {}
            r += 1
            continue
            
        # Check if this row is the header row
        col_4_val = ws.cell(r, 4).value
        if col_4_val and any(h in str(col_4_val) for h in ('ORDER ID', 'លេខបុង')):
            # Row r is headers, Row r+1 is days. We can pass a mock class/ws slice to get_date_columns_map
            # Or we can just parse it manually:
            r_months = r
            r_days = r + 1
            current_month = None
            for c in range(1, ws.max_column + 1):
                m_val = ws.cell(r_months, c).value
                if m_val is not None:
                    m_str = str(m_val).strip()
                    if m_str in ("June", "July"):
                        current_month = m_str
                d_val = ws.cell(r_days, c).value
                if d_val is not None and str(d_val).strip().isdigit() and current_month:
                    date_cols[c] = (current_month, int(str(d_val).strip()))
            r = r_days + 1
            continue
            
        # Count urgent order rows in current_section
        if current_section and date_cols:
            order_id = ws.cell(r, 4).value
            if order_id is not None and str(order_id).strip().isdigit():
                created_date = None
                for c_idx, (m, d) in date_cols.items():
                    cell_val = ws.cell(r, c_idx).value
                    if cell_val == 1 or str(cell_val).strip() == '1':
                        created_date = (m, d)
                        break
                if created_date and is_urgent(created_date[0], created_date[1]):
                    counts[current_section] += 1
                    
        r += 1
    return counts

morning_urgent = get_morning_urgent()
afternoon_urgent = get_afternoon_urgent()

print("Corrected Morning Urgent (9 AM):", morning_urgent)
print("Corrected Afternoon Urgent (2 PM):", afternoon_urgent)

# Write output to comparison_urgent_summary.txt
with open("c:/Users/DELL/Desktop/daily_push/scratch/urgent_comparison_summary.txt", "w", encoding="utf-8") as f:
    f.write("=== URGENT (OVERDUE >= 2 DAYS) REPORT COMPARISON ===\n\n")
    f.write(f"Morning File: {file_morning}\n")
    f.write(f"Afternoon File: {file_afternoon}\n\n")
    f.write(f"| Category | Morning (9 AM) | Afternoon (2 PM) | Change | Change % |\n")
    f.write(f"| :--- | :---: | :---: | :---: | :---: |\n")
    for k in ['Pickup', 'Delivery', 'Pending']:
        m = morning_urgent.get(k, 0)
        a = afternoon_urgent.get(k, 0)
        diff = a - m
        pct = (diff / m * 100) if m > 0 else 0.0
        f.write(f"| {k} | {m} | {a} | {diff:+} | {pct:+.2f}% |\n")
        
    m_tot = sum(morning_urgent.values())
    a_tot = sum(afternoon_urgent.values())
    tot_diff = a_tot - m_tot
    tot_pct = (tot_diff / m_tot * 100) if m_tot > 0 else 0.0
    f.write(f"| **Total Urgent** | **{m_tot}** | **{a_tot}** | **{tot_diff:+}** | **{tot_pct:+.2f}%** |\n")

print("SUCCESS")
