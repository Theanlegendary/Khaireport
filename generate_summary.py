"""
generate_summary.py
Builds a single summary image showing totals per handle per report type.
No order detail — just counts.
"""

import io
from datetime import datetime
from PIL import Image, ImageDraw, ImageFont

# ── Visual config ──────────────────────────────────────────────────────────────
SCALE       = 3  # High-definition scaling factor
FONT_SIZE   = 11 * SCALE
ROW_H       = 24 * SCALE
PAD_X       = 12 * SCALE
PAD_Y       = 6 * SCALE

# Colours
C_TITLE_BG  = (15,  23,  42)   # #0F172A dark navy
C_TITLE_FG  = (255, 255, 255)
C_HEADER_BG = (30,  41,  59)   # #1E293B dark slate
C_HEADER_FG = (255, 255, 255)
C_ROW_BG    = (255, 255, 255)
C_ROW_ALT   = (248, 250, 252)  # #F8FAFC light stripe
C_TOTAL_BG  = (241, 245, 249)  # #F1F5F9
C_TOTAL_FG  = (239, 68,  68)   # #EF4444 red
C_NUM_FG    = (239, 68,  68)   # red numbers
C_TEXT_FG   = (15,  23,  42)
C_BORDER    = (203, 213, 225)  # #CBD5E1

REPORT_TYPES = ['Pickup', 'Delivery', 'Pending']


def _load_font(size, bold=False):
    for name in (['arialbd.ttf', 'Arial Bold.ttf', 'DejaVuSans-Bold.ttf'] if bold
                 else ['arial.ttf', 'Arial.ttf', 'DejaVuSans.ttf']):
        try:
            return ImageFont.truetype(name, size)
        except Exception:
            pass
    return ImageFont.load_default()


def _text_w(draw, text, font):
    try:
        bb = draw.textbbox((0, 0), text, font=font)
        return bb[2] - bb[0]
    except Exception:
        return len(text) * (FONT_SIZE - 1 * SCALE)


def build_summary_image(handle_results: list, overall: dict, today: datetime = None) -> io.BytesIO:
    """
    handle_results: list of dicts with keys 'handle', 'handle_counts'
    overall: dict {'Pickup': n, 'Delivery': n, 'Pending': n}
    Returns BytesIO PNG.
    """
    now = today or datetime.now()
    n_data_rows = len(handle_results)
    
    # Dynamically adjust SCALE to keep the summary image size within Telegram's limits (under 4000px height)
    if n_data_rows > 75:
        scale = 1
    elif n_data_rows > 35:
        scale = 2
    else:
        scale = 3

    font_size   = 11 * scale
    row_h       = 24 * scale
    pad_x       = 12 * scale
    pad_y       = 6 * scale

    fn      = _load_font(font_size, bold=False)
    fn_bold = _load_font(font_size, bold=True)
    fn_sm   = _load_font(font_size - 1 * scale, bold=False)

    # ── Measure column widths ──────────────────────────────────────────────────
    tmp  = Image.new("RGB", (1, 1))
    draw = ImageDraw.Draw(tmp)

    col_headers = ["HANDLE"] + REPORT_TYPES + ["TOTAL"]

    # Handle col width based on max handle name
    handle_strs = [hr['handle'] for hr in handle_results] + ["GRAND TOTAL"]
    w_handle = max(_text_w(draw, s, fn_bold) for s in handle_strs) + pad_x * 2
    w_handle = max(w_handle, 90 * scale)

    # Number columns — fixed width
    w_num = max(_text_w(draw, h, fn_bold) for h in REPORT_TYPES + ["TOTAL"]) + pad_x * 2
    w_num = max(w_num, 72 * scale)

    col_widths = [w_handle] + [w_num] * (len(REPORT_TYPES) + 1)  # +1 for TOTAL
    total_w    = sum(col_widths) + 1

    # ── Row count ──────────────────────────────────────────────────────────────
    n_rows = 1 + 1 + n_data_rows + 1  # title + header + data + grand total
    total_h = n_rows * row_h + 1

    img  = Image.new("RGB", (total_w, total_h), (255, 255, 255))
    draw = ImageDraw.Draw(img)

    def draw_row(row_idx, cells, bg, fg_list=None, bold=False):
        y = row_idx * row_h
        x = 0
        f = fn_bold if bold else fn
        for ci, (cell_text, cw) in enumerate(zip(cells, col_widths)):
            fg = fg_list[ci] if fg_list else C_TEXT_FG
            draw.rectangle([x, y, x + cw, y + row_h], fill=bg)
            if cell_text:
                tw = _text_w(draw, cell_text, f)
                # Center numbers, left-align handle
                if ci == 0:
                    tx = x + pad_x
                else:
                    tx = x + (cw - tw) // 2
                ty = y + (row_h - font_size) // 2
                draw.text((tx, ty), cell_text, font=f, fill=fg)
            draw.rectangle([x, y, x + cw, y + row_h], outline=C_BORDER, width=1 * scale)
            x += cw

    # Title row
    title = f"📊 DAILY TOTAL  {now.strftime('%d/%m/%Y  %H:%M')}"
    draw.rectangle([0, 0, total_w, row_h], fill=C_TITLE_BG)
    f = fn_bold
    tw = _text_w(draw, title, f)
    draw.text((pad_x, (row_h - font_size) // 2), title, font=f, fill=C_TITLE_FG)
    draw.rectangle([0, 0, total_w, row_h], outline=C_BORDER, width=1 * scale)

    # Header row
    draw_row(1, col_headers, C_HEADER_BG,
             fg_list=[C_HEADER_FG] * len(col_headers), bold=True)

    # Data rows
    for i, hr in enumerate(handle_results):
        counts = hr['handle_counts']
        pickup   = counts.get('Pickup',   0)
        delivery = counts.get('Delivery', 0)
        pending  = counts.get('Pending',  0)
        total    = pickup + delivery + pending
        cells = [
            hr['handle'],
            str(pickup)   if pickup   else '',
            str(delivery) if delivery else '',
            str(pending)  if pending  else '',
            str(total),
        ]
        bg = C_ROW_ALT if i % 2 else C_ROW_BG
        fg = [C_TEXT_FG, C_NUM_FG, C_NUM_FG, C_NUM_FG, C_TOTAL_FG]
        draw_row(2 + i, cells, bg, fg_list=fg)

    # Grand Total row
    g_pickup   = overall.get('Pickup',   0)
    g_delivery = overall.get('Delivery', 0)
    g_pending  = overall.get('Pending',  0)
    g_total    = g_pickup + g_delivery + g_pending
    grand_cells = [
        "GRAND TOTAL",
        str(g_pickup),
        str(g_delivery),
        str(g_pending),
        str(g_total),
    ]
    draw_row(2 + n_data_rows, grand_cells, C_TOTAL_BG,
             fg_list=[C_TOTAL_FG] * len(grand_cells), bold=True)

    buf = io.BytesIO()
    img.save(buf, format='PNG', optimize=True)
    buf.seek(0)
    return buf

# ── Total Excel builder ────────────────────────────────────────────────────────

def build_total_excel(result: dict, out_path: str):
    """
    Build a summary Excel with 3 tables (Pickup / Delivery / Pending) on a SINGLE sheet.
    Each table = ALL branches for that type, sorted by POST OFFICE HANDLE.
    Uses raw DataFrames from result['type_data'] — no re-parsing of Excel files.
    """
    import pandas as pd
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
    from openpyxl.utils import get_column_letter

    REPORT_ORDER = ['Pickup', 'Delivery', 'Pending']

    # Index cols per report type
    REPORT_COLS = {
        'Pickup':   ['ZONE', 'POST OFFICE HANDLE', 'CURRENT POST OFFICE', 'ORDER ID', 'Cus name', 'Phone'],
        'Delivery': ['ZONE', 'POST OFFICE HANDLE', 'CURRENT POST OFFICE', 'ORDER ID', 'RECEIVER'],
        'Pending':  ['ZONE', 'POST OFFICE HANDLE', 'CURRENT POST OFFICE', 'ORDER ID'],
    }

    type_data   = result.get('type_data', {})
    day_cols    = result.get('day_cols', [])
    date_col    = result.get('cur_time_col') or 'CURRENT TIME'
    now_str     = datetime.now().strftime('%d.%m_%Hh%M')

    fn    = 'Aptos Narrow'
    RED   = 'EF4444'
    NAVY  = '0F172A'
    SLATE = '1E293B'
    thin  = Side(style='thin', color='BFBFBF')
    bdr   = Border(left=thin, right=thin, top=thin, bottom=thin)

    def _hcell(ws, r, c, val):
        cell = ws.cell(r, c, val)
        cell.font      = Font(name=fn, color='FFFFFF', bold=True, size=10)
        cell.fill      = PatternFill(start_color=SLATE, end_color=SLATE, fill_type='solid')
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border    = bdr

    def _tcell(ws, r, c, val, n_cols):
        cell = ws.cell(r, c, val)
        cell.font      = Font(name=fn, color='FFFFFF', bold=True, size=12)
        cell.fill      = PatternFill(start_color=NAVY, end_color=NAVY, fill_type='solid')
        cell.alignment = Alignment(horizontal='left', vertical='center')
        cell.border    = bdr
        ws.row_dimensions[r].height = 22
        if n_cols > 1:
            ws.merge_cells(start_row=r, end_row=r, start_column=c, end_column=c + n_cols - 1)

    wb = Workbook()
    ws = wb.active
    ws.title = "Report_All"
    
    current_row = 1

    for rn in REPORT_ORDER:
        df = type_data.get(rn)
        # Ensure a DataFrame exists; if missing or empty, create an empty one with required columns
        if df is None:
            df = pd.DataFrame(columns=REPORT_COLS[rn])
        elif df.empty:
            for col in REPORT_COLS[rn]:
                if col not in df.columns:
                    df[col] = ''

        idx_cols = REPORT_COLS[rn]

        # Ensure all index cols exist
        for col in idx_cols:
            if col not in df.columns:
                df[col] = ''

        # Extract day from CURRENT TIME
        if date_col in df.columns:
            parsed = pd.to_datetime(df[date_col], dayfirst=True, format='mixed', errors='coerce')
            df = df.copy()
            df['_day'] = parsed.dt.strftime('%d').fillna('')
        else:
            df = df.copy()
            df['_day'] = ''

        # Only keep day_cols that have data
        days_present = set(df['_day'].unique()) - {'', 'NaT', 'nan'}
        active_days  = [d for d in day_cols if d in days_present]

        # Build day columns
        for d in active_days:
            df[d] = (df['_day'] == d).astype(int)
        df['Grand Total'] = 1

        agg = df.groupby(idx_cols, sort=False, dropna=False)[
            active_days + ['Grand Total']
        ].sum().reset_index()

        # Replace 0 with '' for display
        for d in active_days:
            agg[d] = agg[d].apply(lambda v: int(v) if v > 0 else '')

        # Sort by POST OFFICE HANDLE then CURRENT POST OFFICE
        sort_cols = [c for c in ['POST OFFICE HANDLE', 'CURRENT POST OFFICE', 'ORDER ID']
                     if c in agg.columns]
        agg = agg.sort_values(sort_cols).reset_index(drop=True)

        all_cols = idx_cols + active_days + ['Grand Total']
        n        = len(all_cols)

        # Title
        _tcell(ws, current_row, 1, f"{rn.upper()} BILL CHECK  {now_str}  — ALL BRANCHES", n)

        # Headers
        ws.row_dimensions[current_row + 1].height = 17
        for ci, col in enumerate(all_cols, start=1):
            _hcell(ws, current_row + 1, ci, col)

        # Build order creation map from df (original DataFrame)
        order_created_map = {}
        if 'ORDER ID' in df.columns and 'CREATED DATE' in df.columns:
            parsed_created = pd.to_datetime(df['CREATED DATE'], dayfirst=True, format='mixed', errors='coerce')
            for order_id, dt in zip(df['ORDER ID'].astype(str).str.strip(), parsed_created):
                if pd.notna(dt):
                    order_created_map[order_id] = dt.date()

        # Data rows
        day_totals  = {d: 0 for d in active_days}
        grand_total = 0

        for ri, row in agg.iterrows():
            r = current_row + 2 + ri
            ws.row_dimensions[r].height = 15
            gt_val = int(row.get('Grand Total', 0))
            grand_total += gt_val

            # Check if this order is over 1 day old (matching generate_report.py HIGHLIGHT_OVER_DAYS = 1)
            is_overdue = False
            if 'ORDER ID' in row:
                order_id = str(row['ORDER ID']).strip()
                if order_id in order_created_map:
                    created_date = order_created_map[order_id]
                    delta = datetime.now().date() - created_date
                    if delta.days > 1:
                        is_overdue = True

            for ci, col in enumerate(all_cols, start=1):
                val  = row.get(col, '')
                cell = ws.cell(r, ci, val if val != '' else None)
                cell.border = bdr
                cell.font   = Font(name=fn, size=10)

                # Fill with soft red if overdue
                if is_overdue:
                    cell.fill = PatternFill(start_color='FFEBEB', end_color='FFEBEB', fill_type='solid')

                if col in active_days:
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    if isinstance(val, (int, float)) and val:
                        day_totals[col] = day_totals.get(col, 0) + int(val)
                elif col == 'Grand Total':
                    cell.font      = Font(name=fn, color=RED, bold=True, size=10)
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                else:
                    cell.alignment = Alignment(horizontal='left', vertical='center')

        # Grand Total footer
        gt_row = current_row + 2 + len(agg)
        ws.row_dimensions[gt_row].height = 17
        for ci, col in enumerate(all_cols, start=1):
            cell = ws.cell(gt_row, ci)
            cell.font      = Font(name=fn, color=RED, bold=True, size=10)
            cell.fill      = PatternFill(start_color='F1F5F9', end_color='F1F5F9', fill_type='solid')
            cell.border    = bdr
            cell.alignment = Alignment(horizontal='center', vertical='center')
            if ci == 1:
                cell.value = 'Grand Total'
            elif col in active_days:
                cell.value = day_totals.get(col) or None
            elif col == 'Grand Total':
                cell.value = grand_total or None

        # Column widths (update max width over all tables)
        for ci, col in enumerate(all_cols, start=1):
            letter = get_column_letter(ci)
            if col.isdigit() and len(col) == 2:
                ws.column_dimensions[letter].width = 5
            elif col == 'Grand Total':
                ws.column_dimensions[letter].width = 12
            elif col == 'ZONE':
                ws.column_dimensions[letter].width = 9
            elif col in ('Cus name', 'RECEIVER'):
                max_len = max(
                    (len(str(ws.cell(r_iter, ci).value or ''))
                     for r_iter in range(current_row, gt_row + 1)),
                    default=20
                )
                existing = ws.column_dimensions[letter].width
                ws.column_dimensions[letter].width = max(existing if existing else 0, min(max(max_len + 3, 22), 50))
            elif col == 'Phone':
                max_len = max(
                    (len(str(ws.cell(r_iter, ci).value or ''))
                     for r_iter in range(current_row, gt_row + 1)),
                    default=14
                )
                existing = ws.column_dimensions[letter].width
                ws.column_dimensions[letter].width = max(existing if existing else 0, min(max(max_len + 3, 16), 35))
            else:
                max_len = max(
                    (len(str(ws.cell(r_iter, ci).value or ''))
                     for r_iter in range(current_row, gt_row + 1)),
                    default=8
                )
                existing = ws.column_dimensions[letter].width
                ws.column_dimensions[letter].width = max(existing if existing else 0, min(max(max_len + 2, 10), 28))

        # Advance current_row for the next table
        current_row = gt_row + 3

    wb.save(out_path)
    return out_path
