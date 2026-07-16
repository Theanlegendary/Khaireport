"""Add run_mega_combined to pivot.py that builds zone table + mega table in one Excel."""
import os

HERE = os.path.dirname(os.path.abspath(__file__))
path = os.path.join(HERE, 'pivot.py')

with open(path, 'r', encoding='utf-8') as f:
    content = f.read()

# Add new function before the if __name__ block
NEW_FUNC = '''

def run_mega_combined(source_path, out_path, config):
    """Build combined Excel: Zone pivot on top + MEGA pivot below."""
    rows = read_source(source_path)
    zone_cfg = config.get("zone_mapping", {})
    pivot_cfg = config.get("pivot", {})

    # Build zone pivot (all pending statuses by zone)
    zone_tree, zone_day_keys, zone_month = build_pivot(rows, pivot_cfg, zone_cfg)

    # Build mega pivot
    mega_tree, mega_day_keys, mega_urgent = build_mega_pivot(rows, pivot_cfg, zone_cfg)

    # Combined day keys
    all_day_keys = sorted(set(zone_day_keys + mega_day_keys))

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "MEGA + Zone"

    hdr_fill = PatternFill("solid", fgColor="BDD7EE")
    hdr_font = Font(name="Calibri", size=11, bold=True, color="000000")
    total_font = Font(name="Calibri", size=11, bold=True, color="000000")
    urgent_font = Font(name="Calibri", size=11, bold=True, color="FF0000")
    data_font = Font(name="Calibri", size=11, color="000000")
    title_font = Font(name="Calibri", size=11, bold=True, color="FF0000")

    # ── ZONE TABLE (top) ──────────────────────────────────────────────
    r = 1
    title_time = datetime.now().strftime("%d.%m_%HH%M")
    ws.cell(r, 1, f"BILLS {title_time}").font = title_font
    ws.row_dimensions[r].height = 20
    r += 1

    # Zone header row
    ws.cell(r, 1, "Zone").font = hdr_font
    ws.cell(r, 1).fill = hdr_fill
    ws.cell(r, 1).border = _BORDER
    col_idx_map = {}
    for idx, dk in enumerate(all_day_keys):
        col_num = 2 + idx
        col_idx_map[dk] = col_num
        cell = ws.cell(r, col_num, f"{dk[1]:02d}")
        cell.font = hdr_font
        cell.fill = hdr_fill
        cell.border = _BORDER
        cell.alignment = _CENTER
    total_col = 2 + len(all_day_keys)
    ws.cell(r, total_col, "Total").font = hdr_font
    ws.cell(r, total_col).fill = hdr_fill
    ws.cell(r, total_col).border = _BORDER
    ws.cell(r, total_col).alignment = _CENTER
    r += 1

    # Zone data rows
    zone_col_totals = defaultdict(int)
    zone_grand = 0
    for zone_name in sorted(zone_tree.keys()):
        ws.cell(r, 1, zone_name).font = data_font
        ws.cell(r, 1).border = _BORDER
        row_sum = 0
        for po in zone_tree[zone_name]:
            for oid, dk in zone_tree[zone_name][po].items():
                if dk in col_idx_map:
                    col_num = col_idx_map[dk]
                    cur = ws.cell(r, col_num).value or 0
                    ws.cell(r, col_num, cur + 1)
                    zone_col_totals[dk] += 1
                    row_sum += 1
        for c in range(2, total_col + 1):
            cell = ws.cell(r, c)
            cell.border = _BORDER
            cell.alignment = _CENTER
            cell.font = data_font
            if cell.value == 0 or cell.value is None:
                cell.value = None
        ws.cell(r, total_col, row_sum).font = data_font
        ws.cell(r, total_col).border = _BORDER
        ws.cell(r, total_col).alignment = _CENTER
        zone_grand += row_sum
        ws.row_dimensions[r].height = 20
        r += 1

    # Zone total row
    ws.cell(r, 1, "Total").font = total_font
    ws.cell(r, 1).fill = hdr_fill
    ws.cell(r, 1).border = _BORDER
    for dk, col_num in col_idx_map.items():
        val = zone_col_totals.get(dk, 0)
        cell = ws.cell(r, col_num)
        cell.value = val if val > 0 else None
        cell.font = total_font
        cell.fill = hdr_fill
        cell.border = _BORDER
        cell.alignment = _CENTER
    ws.cell(r, total_col, zone_grand).font = total_font
    ws.cell(r, total_col).fill = hdr_fill
    ws.cell(r, total_col).border = _BORDER
    ws.cell(r, total_col).alignment = _CENTER
    ws.row_dimensions[r].height = 20
    r += 2  # gap

    # ── MEGA TABLE (bottom) ───────────────────────────────────────────
    ws.cell(r, 1, f"TỒN MEGA CHECK {title_time}").font = title_font
    ws.row_dimensions[r].height = 20
    r += 1

    # Mega header
    ws.cell(r, 1, "Row Labels").font = hdr_font
    ws.cell(r, 1).fill = hdr_fill
    ws.cell(r, 1).border = _BORDER
    for dk, col_num in col_idx_map.items():
        cell = ws.cell(r, col_num, f"{dk[1]:02d}")
        cell.font = hdr_font
        cell.fill = hdr_fill
        cell.border = _BORDER
        cell.alignment = _CENTER
    ws.cell(r, total_col, "Total").font = hdr_font
    ws.cell(r, total_col).fill = hdr_fill
    ws.cell(r, total_col).border = _BORDER
    ws.cell(r, total_col).alignment = _CENTER
    urgent_col = total_col + 1
    ws.cell(r, urgent_col, "Urgent").font = urgent_font
    ws.cell(r, urgent_col).fill = hdr_fill
    ws.cell(r, urgent_col).border = _BORDER
    ws.cell(r, urgent_col).alignment = _CENTER
    r += 1

    # Mega data rows
    mega_col_totals = defaultdict(int)
    mega_grand = 0
    mega_urgent_total = 0
    for hub in sorted(mega_tree.keys()):
        ws.cell(r, 1, hub).font = data_font
        ws.cell(r, 1).border = _BORDER
        row_sum = 0
        for dk in all_day_keys:
            col_num = col_idx_map[dk]
            val = mega_tree[hub].get(dk, 0)
            cell = ws.cell(r, col_num)
            cell.border = _BORDER
            cell.alignment = _CENTER
            cell.font = data_font
            if val > 0:
                cell.value = val
                row_sum += val
                mega_col_totals[dk] += val
        ws.cell(r, total_col, row_sum).font = data_font
        ws.cell(r, total_col).border = _BORDER
        ws.cell(r, total_col).alignment = _CENTER
        # Urgent
        hub_urgent = mega_urgent.get(hub, 0)
        urg_cell = ws.cell(r, urgent_col)
        urg_cell.border = _BORDER
        urg_cell.alignment = _CENTER
        if hub_urgent > 0:
            urg_cell.value = hub_urgent
            urg_cell.font = urgent_font
        else:
            urg_cell.font = data_font
        mega_urgent_total += hub_urgent
        mega_grand += row_sum
        ws.row_dimensions[r].height = 20
        r += 1

    # Mega total row
    ws.cell(r, 1, "Total").font = total_font
    ws.cell(r, 1).fill = hdr_fill
    ws.cell(r, 1).border = _BORDER
    for dk, col_num in col_idx_map.items():
        val = mega_col_totals.get(dk, 0)
        cell = ws.cell(r, col_num)
        cell.value = val if val > 0 else None
        cell.font = total_font
        cell.fill = hdr_fill
        cell.border = _BORDER
        cell.alignment = _CENTER
    ws.cell(r, total_col, mega_grand).font = urgent_font
    ws.cell(r, total_col).fill = hdr_fill
    ws.cell(r, total_col).border = _BORDER
    ws.cell(r, total_col).alignment = _CENTER
    ws.cell(r, urgent_col, mega_urgent_total).font = urgent_font
    ws.cell(r, urgent_col).fill = hdr_fill
    ws.cell(r, urgent_col).border = _BORDER
    ws.cell(r, urgent_col).alignment = _CENTER

    # Column widths
    ws.column_dimensions["A"].width = 16
    for c in range(2, urgent_col + 1):
        ws.column_dimensions[get_column_letter(c)].width = 7
    ws.column_dimensions[get_column_letter(urgent_col)].width = 9

    wb.save(out_path)
    return out_path, mega_grand

'''

# Insert before if __name__
marker = 'if __name__ == "__main__":'
if 'def run_mega_combined' not in content:
    content = content.replace(marker, NEW_FUNC + '\n' + marker)
    with open(path, 'w', encoding='utf-8') as f:
        f.write(content)
    print("DONE: added run_mega_combined")
else:
    print("SKIP: already exists")
