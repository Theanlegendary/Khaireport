"""
generate_summary.py
Builds a zone push summary image with the layout:
  Row 1  : Title bar  — "DAILY REPORT — ZONE3  06/07/2026  09:30"
  Row 2  : Month sub-header spanning date columns
  Row 3  : Column headers — HANDLE | Pickup | Delivery | Pending | <DD> … | TOTAL | URGENT
  Row 4+ : Data — branch row with counts; URGENT column red bold
  Last   : Grand Total — all red bold
"""

import io
import calendar as _calendar
from datetime import datetime, date as _date
from PIL import Image, ImageDraw, ImageFont

SUMMARY_HEADER_KHMER = {
    "HANDLE": "ប៉ុស្តិ៍ / ហាង",
    "Pickup": "ត្រូវយក",
    "Delivery": "ត្រូវដឹក",
    "Pending": "កំពុងរង់ចាំ",
    "TOTAL": "សរុប",
    "URGENT": "ប្រញាប់",
    "U.Pickup": "ប្រញាប់.យក",
    "U.Delivery": "ប្រញាប់.ដឹក",
    "U.Pending": "ប្រញាប់.រង់ចាំ",
    "GRAND TOTAL": "សរុបទាំងអស់"
}

# ── Palette ────────────────────────────────────────────────────────────────────
C_TITLE_BG    = ( 10,  15,  35)   # near-black navy
C_TITLE_FG    = (255, 255, 255)
C_MONTH_BG    = ( 22,  34,  64)   # deep indigo for month row
C_MONTH_FG    = (180, 200, 255)   # soft blue-white
C_HEADER_BG   = ( 30,  45,  80)   # dark slate-blue
C_HEADER_FG   = (255, 255, 255)
C_URGENT_HDR  = (180,  20,  20)   # dark red header for URGENT col
C_ROW_BG      = (255, 255, 255)
C_ROW_ALT     = (245, 248, 255)   # very light blue stripe
C_TOTAL_BG    = (232, 238, 250)   # soft blue-grey footer
C_TOTAL_FG    = (200,  30,  30)   # strong red
C_NUM_FG      = ( 30,  30, 160)   # deep blue for Pickup/Delivery/Pending counts
C_PENDING_FG  = (180,  80,   0)   # amber for Pending counts
C_DATE_FG     = ( 30,  30, 160)   # blue for date counts
C_URGENT_FG   = (210,  30,  30)   # red for urgent counts
C_HANDLE_FG   = ( 10,  15,  40)   # near-black for branch name
C_BORDER      = (180, 195, 220)
C_BORDER_DARK = ( 80, 100, 140)   # darker border for section separators

_WIN_FONTS = "C:/Windows/Fonts"


def _load_font(size, bold=False):
    candidates = (
        [
            f"{_WIN_FONTS}/calibrib.ttf",
            f"{_WIN_FONTS}/arialbd.ttf",
            f"{_WIN_FONTS}/verdanab.ttf",
            f"{_WIN_FONTS}/DejaVuSans-Bold.ttf",
            "arialbd.ttf", "DejaVuSans-Bold.ttf",
        ]
        if bold else
        [
            f"{_WIN_FONTS}/calibri.ttf",
            f"{_WIN_FONTS}/arial.ttf",
            f"{_WIN_FONTS}/verdana.ttf",
            f"{_WIN_FONTS}/DejaVuSans.ttf",
            "arial.ttf", "DejaVuSans.ttf",
        ]
    )
    for name in candidates:
        try:
            return ImageFont.truetype(name, size)
        except Exception:
            pass
    return ImageFont.load_default()


def _tw(draw, text, font):
    try:
        bb = draw.textbbox((0, 0), text, font=font)
        return bb[2] - bb[0]
    except Exception:
        return len(text) * max(8, font.size - 2)


def _th(draw, text, font):
    try:
        bb = draw.textbbox((0, 0), text, font=font)
        return bb[3] - bb[1]
    except Exception:
        return font.size


def _draw_cell(draw, x, y, w, h, bg, text, font, fg, align="center", pad=8,
               border=True, border_col=None, border_w=1):
    draw.rectangle([x, y, x + w - 1, y + h - 1], fill=bg)
    if text:
        tw = _tw(draw, text, font)
        th = _th(draw, text, font)
        ty = y + (h - th) // 2
        if align == "center":
            tx = x + (w - tw) // 2
        elif align == "right":
            tx = x + w - tw - pad
        else:
            tx = x + pad
        draw.text((tx, ty), text, font=font, fill=fg)
    if border:
        bc = border_col or C_BORDER
        draw.rectangle([x, y, x + w - 1, y + h - 1], outline=bc, width=border_w)


def build_summary_image(
    handle_results: list,
    overall: dict,
    today: datetime = None,
    zone_label: str = "",
    day_date_counts: dict = None,   # {handle: {date_obj: count}} for date columns
    urgent_counts: dict = None,     # {handle: urgent_count}
) -> io.BytesIO:
    """
    New layout:
      Row 1  : Title   — "DAILY REPORT — ZONE3  06/07/2026 09:30"
      Row 2  : Month sub-header spanning date columns (blank over fixed cols)
      Row 3  : Headers — HANDLE | Pickup | Delivery | Pending | DD DD DD … | TOTAL | URGENT
      Row 4+ : Data rows
      Last   : Grand Total (all red)

    day_date_counts: optional {handle: {date: int}} — if provided, adds date columns.
    urgent_counts:   optional {handle: int}          — if provided, adds URGENT column.
    """
    now = today or datetime.now()
    n_data_rows = len(handle_results)

    # ── Scale factor ───────────────────────────────────────────────────────────
    if n_data_rows > 75:
        sc = 1
    elif n_data_rows > 35:
        sc = 2
    else:
        sc = 3

    FS       = 11 * sc   # base font size
    FS_TITLE = 13 * sc   # title font size
    FS_SM    = 9  * sc   # small (month row)
    ROW_H    = 26 * sc
    TITLE_H  = 34 * sc
    MONTH_H  = 18 * sc
    PAD      =  8 * sc

    fn       = _load_font(FS,       bold=False)
    fn_b     = _load_font(FS,       bold=True)
    fn_title = _load_font(FS_TITLE, bold=True)
    fn_sm    = _load_font(FS_SM,    bold=True)

    # ── Collect all unique sorted dates ───────────────────────────────────────
    all_dates: list[_date] = []
    if day_date_counts:
        date_set = set()
        for dc in day_date_counts.values():
            date_set.update(dc.keys())
        all_dates = sorted(date_set)

    # ── Column definitions ────────────────────────────────────────────────────
    tmp  = Image.new("RGB", (1, 1))
    draw = ImageDraw.Draw(tmp)

    handle_strs = [hr["handle"] for hr in handle_results] + ["GRAND TOTAL"]
    W_HANDLE = max(_tw(draw, s, fn_b) for s in handle_strs) + PAD * 2
    W_HANDLE = max(W_HANDLE, 80 * sc)

    W_NUM    = max(_tw(draw, h, fn_b) for h in ["Pickup", "Delivery", "Pending", "TOTAL"]) + PAD * 2
    W_NUM    = max(W_NUM, 56 * sc)

    W_DATE   = max(_tw(draw, "00", fn_b) + PAD * 2, 32 * sc)
    W_URGENT = max(_tw(draw, "URGENT", fn_sm) + PAD * 2, 48 * sc)
    W_U_COL  = max(_tw(draw, "U.Delivery", fn_sm) + PAD * 2, 48 * sc)

    # Column order: Handle | Pickup | Delivery | Pending | [dates…] | TOTAL | URGENT
    fixed_cols  = ["HANDLE", "Pickup", "Delivery", "Pending"]
    date_labels = [f"{d.day:02d}" for d in all_dates]
    tail_cols   = ["TOTAL"]
    
    has_split_urgent = False
    if urgent_counts:
        first_val = next(iter(urgent_counts.values()))
        if isinstance(first_val, dict):
            has_split_urgent = True

    if urgent_counts is not None:
        if has_split_urgent:
            tail_cols.extend(["U.Pickup", "U.Delivery", "U.Pending", "URGENT"])
        else:
            tail_cols.append("URGENT")

    col_widths = (
        [W_HANDLE, W_NUM, W_NUM, W_NUM]
        + [W_DATE] * len(all_dates)
        + [W_NUM]                          # TOTAL
    )
    if urgent_counts is not None:
        if has_split_urgent:
            col_widths.extend([W_U_COL, W_U_COL, W_U_COL, W_URGENT])
        else:
            col_widths.append(W_URGENT)
    col_labels = fixed_cols + date_labels + tail_cols

    n_cols  = len(col_widths)
    total_w = sum(col_widths) + 1

    # ── Canvas height ─────────────────────────────────────────────────────────
    # Title + month row + header row + data rows + grand total
    total_h = TITLE_H + MONTH_H + ROW_H + n_data_rows * ROW_H + ROW_H + 1

    img  = Image.new("RGB", (total_w, total_h), C_ROW_BG)
    draw = ImageDraw.Draw(img)

    # ── Helper: draw a full row of cells ──────────────────────────────────────
    def _row(y, h, cells, bgs, fgs, fonts, aligns, border_col=C_BORDER):
        x = 0
        for ci, (text, cw) in enumerate(zip(cells, col_widths)):
            _draw_cell(draw, x, y, cw, h,
                       bg=bgs[ci] if isinstance(bgs, list) else bgs,
                       text=text,
                       font=fonts[ci] if isinstance(fonts, list) else fonts,
                       fg=fgs[ci] if isinstance(fgs, list) else fgs,
                       align=aligns[ci] if isinstance(aligns, list) else aligns,
                       pad=PAD, border=True, border_col=border_col)
            x += cw

    # ── Row 1: Title ──────────────────────────────────────────────────────────
    y = 0
    zone_part = f" — {zone_label.upper()}" if zone_label else ""
    title_text = f"DAILY REPORT{zone_part}    {now.strftime('%d/%m/%Y  %H:%M')}"
    draw.rectangle([0, 0, total_w - 1, TITLE_H - 1], fill=C_TITLE_BG)
    tw = _tw(draw, title_text, fn_title)
    th = _th(draw, title_text, fn_title)
    draw.text((PAD * 2, (TITLE_H - th) // 2), title_text, font=fn_title, fill=C_TITLE_FG)
    # right-side date accent bar
    accent_w = 6 * sc
    draw.rectangle([total_w - accent_w - 1, 0, total_w - 1, TITLE_H - 1], fill=(60, 100, 200))
    draw.rectangle([0, 0, total_w - 1, TITLE_H - 1], outline=C_BORDER_DARK, width=sc)
    y += TITLE_H

    # ── Row 2: Month sub-header ───────────────────────────────────────────────
    # Fixed cols + tail cols are blank (same dark bg); date cols show month name merged
    month_cells  = [""] * n_cols
    month_bgs    = [C_MONTH_BG] * n_cols
    month_fgs    = [C_MONTH_FG] * n_cols
    month_fonts  = [fn_sm] * n_cols
    month_aligns = ["center"] * n_cols

    # Group consecutive dates by month and write month name into first cell of group
    if all_dates:
        n_fixed = 4   # HANDLE + 3 type cols
        groups = []
        cur_mo, g_start = None, None
        for i, d in enumerate(all_dates):
            mo = (d.year, d.month)
            if mo != cur_mo:
                if cur_mo is not None:
                    groups.append((cur_mo, g_start, i - 1))
                cur_mo, g_start = mo, i
        if cur_mo is not None:
            groups.append((cur_mo, g_start, len(all_dates) - 1))

        for (yr, mo), gi_start, gi_end in groups:
            col_idx = n_fixed + gi_start   # 0-based column index
            month_cells[col_idx] = _calendar.month_abbr[mo].upper()

    # Draw month row manually (need to visually merge month spans)
    x = 0
    for ci, cw in enumerate(col_widths):
        is_date_col = (4 <= ci < 4 + len(all_dates))
        bg = C_MONTH_BG
        # lighter bg for non-date cols
        cell_bg = C_MONTH_BG if is_date_col else C_HEADER_BG
        _draw_cell(draw, x, y, cw, MONTH_H,
                   bg=cell_bg,
                   text=month_cells[ci],
                   font=fn_sm, fg=C_MONTH_FG,
                   align="center", pad=PAD,
                   border=True, border_col=C_BORDER_DARK)
        x += cw
    y += MONTH_H

    # ── Row 3: Column headers ─────────────────────────────────────────────────
    hdr_bgs = []
    hdr_fgs = []
    for ci, label in enumerate(col_labels):
        if label == "URGENT":
            hdr_bgs.append(C_URGENT_HDR)
            hdr_fgs.append((255, 230, 230))
        elif label in ("TOTAL",):
            hdr_bgs.append((20, 60, 120))
            hdr_fgs.append(C_HEADER_FG)
        else:
            hdr_bgs.append(C_HEADER_BG)
            hdr_fgs.append(C_HEADER_FG)

    _row(y, ROW_H,
         cells=col_labels,
         bgs=hdr_bgs,
         fgs=hdr_fgs,
         fonts=fn_b,
         aligns=["left"] + ["center"] * (n_cols - 1),
         border_col=C_BORDER_DARK)
    y += ROW_H

    # ── Data rows ─────────────────────────────────────────────────────────────
    for i, hr in enumerate(handle_results):
        counts   = hr["handle_counts"]
        handle   = hr["handle"]
        pickup   = counts.get("Pickup",   0)
        delivery = counts.get("Delivery", 0)
        pending  = counts.get("Pending",  0)
        total    = pickup + delivery + pending
        urgent   = (urgent_counts or {}).get(handle, 0)

        row_bg = C_ROW_ALT if i % 2 else C_ROW_BG

        cells  = [handle,
                  str(pickup)   if pickup   else "",
                  str(delivery) if delivery else "",
                  str(pending)  if pending  else ""]

        fgs    = [C_HANDLE_FG,
                  C_NUM_FG,
                  C_NUM_FG,
                  C_PENDING_FG]

        fonts  = [fn_b, fn, fn, fn]
        aligns = ["left", "center", "center", "center"]

        # Date columns
        day_dc = (day_date_counts or {}).get(handle, {})
        for d in all_dates:
            cnt = day_dc.get(d, 0)
            cells.append(str(cnt) if cnt else "")
            fgs.append(C_DATE_FG)
            fonts.append(fn)
            aligns.append("center")

        # TOTAL
        cells.append(str(total) if total else "")
        fgs.append(C_TOTAL_FG)
        fonts.append(fn_b)
        aligns.append("center")

        # URGENT
        urgent = 0
        if urgent_counts:
            val_u = urgent_counts.get(hr["handle"], 0)
            if isinstance(val_u, dict):
                urgent = sum(val_u.values())
            else:
                urgent = val_u

        if urgent_counts is not None:
            if has_split_urgent:
                h_urg = urgent_counts.get(hr["handle"], {})
                u_p = h_urg.get("Pickup", 0)
                u_d = h_urg.get("Delivery", 0)
                u_pen = h_urg.get("Pending", 0)
                u_tot = u_p + u_d + u_pen
                
                cells.extend([
                    str(u_p) if u_p else "",
                    str(u_d) if u_d else "",
                    str(u_pen) if u_pen else "",
                    str(u_tot) if u_tot else ""
                ])
                fgs.extend([C_URGENT_FG] * 4)
                fonts.extend([fn_b] * 4)
                aligns.extend(["center"] * 4)
            else:
                cells.append(str(urgent) if urgent else "")
                fgs.append(C_URGENT_FG)
                fonts.append(fn_b)
                aligns.append("center")

        bgs = [row_bg] * n_cols
        # Tint urgent cell if non-zero
        if urgent_counts is not None:
            if has_split_urgent:
                h_urg = urgent_counts.get(hr["handle"], {})
                u_p = h_urg.get("Pickup", 0)
                u_d = h_urg.get("Delivery", 0)
                u_pen = h_urg.get("Pending", 0)
                u_tot = u_p + u_d + u_pen
                if u_tot > 0:
                    bgs[-4:] = [(255, 235, 235)] * 4
            elif urgent > 0:
                bgs[-1] = (255, 235, 235)

        _row(y, ROW_H,
             cells=cells, bgs=bgs, fgs=fgs, fonts=fonts, aligns=aligns,
             border_col=C_BORDER)
        y += ROW_H

    # ── Grand Total row ───────────────────────────────────────────────────────
    g_pickup   = overall.get("Pickup",   0)
    g_delivery = overall.get("Delivery", 0)
    g_pending  = overall.get("Pending",  0)
    g_total    = g_pickup + g_delivery + g_pending

    gt_cells  = ["GRAND TOTAL",
                 str(g_pickup)   if g_pickup   else "",
                 str(g_delivery) if g_delivery else "",
                 str(g_pending)  if g_pending  else ""]

    # Date totals
    for d in all_dates:
        day_total = sum(
            (day_date_counts or {}).get(hr["handle"], {}).get(d, 0)
            for hr in handle_results
        )
        gt_cells.append(str(day_total) if day_total else "")

    gt_cells.append(str(g_total) if g_total else "")
    
    if urgent_counts is not None:
        if has_split_urgent:
            g_p = sum(u.get("Pickup", 0) for u in urgent_counts.values() if isinstance(u, dict))
            g_d = sum(u.get("Delivery", 0) for u in urgent_counts.values() if isinstance(u, dict))
            g_pen = sum(u.get("Pending", 0) for u in urgent_counts.values() if isinstance(u, dict))
            g_tot = g_p + g_d + g_pen
            
            gt_cells.extend([
                str(g_p) if g_p else "",
                str(g_d) if g_d else "",
                str(g_pen) if g_pen else "",
                str(g_tot) if g_tot else ""
            ])
        else:
            g_urgent = sum((urgent_counts or {}).values() if not isinstance(next(iter((urgent_counts or {}).values()), 0), dict) else [0])
            gt_cells.append(str(g_urgent) if g_urgent else "")

    gt_bgs = [C_TOTAL_BG] * n_cols
    if urgent_counts is not None:
        if has_split_urgent:
            g_p = sum(u.get("Pickup", 0) for u in urgent_counts.values() if isinstance(u, dict))
            g_d = sum(u.get("Delivery", 0) for u in urgent_counts.values() if isinstance(u, dict))
            g_pen = sum(u.get("Pending", 0) for u in urgent_counts.values() if isinstance(u, dict))
            g_tot = g_p + g_d + g_pen
            if g_tot > 0:
                gt_bgs[-4:] = [(255, 200, 200)] * 4
        else:
            g_urgent = sum((urgent_counts or {}).values() if not isinstance(next(iter((urgent_counts or {}).values()), 0), dict) else [0])
            if g_urgent > 0:
                gt_bgs[-1] = (255, 200, 200)

    _row(y, ROW_H,
         cells=gt_cells,
         bgs=gt_bgs,
         fgs=C_TOTAL_FG,
         fonts=fn_b,
         aligns=["left"] + ["center"] * (n_cols - 1),
         border_col=C_BORDER_DARK)

    # ── Telegram aspect-ratio guard (max 20:1) ────────────────────────────────
    w, h = img.size
    max_ratio = 18.0
    new_w, new_h = w, h
    if h > 0 and w / h > max_ratio:
        new_h = int(w / max_ratio)
    elif w > 0 and h / w > max_ratio:
        new_w = int(h / max_ratio)
    if (new_w, new_h) != (w, h):
        padded = Image.new("RGB", (new_w, new_h), C_ROW_BG)
        padded.paste(img, (0, 0))
        img = padded

    buf = io.BytesIO()
    img.save(buf, format="PNG", optimize=True)
    buf.seek(0)
    return buf


# ── Total Excel builder ────────────────────────────────────────────────────────

def build_total_excel(result: dict, out_path: str):
    """
    Build a summary Excel with 3 tables (Pickup / Delivery / Pending) on a SINGLE sheet.
    """
    import calendar
    import pandas as pd
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
    from openpyxl.utils import get_column_letter

    REPORT_ORDER = ['Pickup', 'Delivery', 'Pending']
    REPORT_COLS = {
        'Pickup':   ['ZONE', 'POST OFFICE HANDLE', 'CURRENT POST OFFICE', 'ORDER ID', 'Cus name', 'Phone'],
        'Delivery': ['ZONE', 'POST OFFICE HANDLE', 'CURRENT POST OFFICE', 'ORDER ID', 'RECEIVER'],
        'Pending':  ['ZONE', 'POST OFFICE HANDLE', 'CURRENT POST OFFICE', 'ORDER ID', 'REMARK', 'NEXT_ACTION'],
    }

    type_data = result.get('type_data', {})
    order_status_map = result.get('order_status_map', {})
    day_cols  = result.get('day_cols', [])
    date_col  = result.get('cur_time_col') or 'CURRENT TIME'
    now_str   = datetime.now().strftime('%d.%m_%Hh%M')

    fn    = 'Segoe UI'
    RED   = 'EF4444'
    NAVY  = '0F172A'
    SLATE = '1E293B'
    thin  = Side(style='thin', color='BFBFBF')
    bdr   = Border(left=thin, right=thin, top=thin, bottom=thin)

    wb = Workbook()
    ws = wb.active
    ws.title = "Report_All"
    current_row = 1

    for rn in REPORT_ORDER:
        df = type_data.get(rn)
        if df is None:
            df = pd.DataFrame(columns=REPORT_COLS[rn])
        elif df.empty:
            df = df.copy()
            for col in REPORT_COLS[rn]:
                if col not in df.columns:
                    df[col] = ''

        idx_cols = REPORT_COLS[rn]
        for col in idx_cols:
            if col not in df.columns:
                df[col] = ''

        if date_col in df.columns:
            parsed = pd.to_datetime(df[date_col], dayfirst=True, format='mixed', errors='coerce')
            df = df.copy()
            df['_date'] = parsed.dt.date
        else:
            df = df.copy()
            df['_date'] = None

        dates_present = set(df['_date'].dropna().unique())
        # Always include today's date so today's column is never missing
        dates_present.add(datetime.now().date())
        active_days = [d for d in day_cols if d in dates_present]

        for d in active_days:
            df[d] = (df['_date'] == d).astype(int)
        df['Grand Total'] = 1

        agg = df.groupby(idx_cols, sort=False, dropna=False)[
            active_days + ['Grand Total']
        ].sum().reset_index()

        for d in active_days:
            agg[d] = agg[d].apply(lambda v: int(v) if v > 0 else '')

        sort_cols = [c for c in ['POST OFFICE HANDLE', 'CURRENT POST OFFICE', 'ORDER ID']
                     if c in agg.columns]
        agg = agg.sort_values(sort_cols).reset_index(drop=True)

        all_cols = idx_cols + active_days + ['Grand Total']
        n_idx    = len(idx_cols)
        n_total  = len(all_cols)

        title_row = current_row
        ws.row_dimensions[title_row].height = 22
        tc = ws.cell(title_row, 1, f"{rn.upper()} BILL CHECK  {now_str}  — ALL BRANCHES")
        tc.font      = Font(name=fn, color='FFFFFF', bold=True, size=12)
        tc.fill      = PatternFill(start_color=NAVY, end_color=NAVY, fill_type='solid')
        tc.alignment = Alignment(horizontal='left', vertical='center')
        tc.border    = bdr
        for ci in range(2, n_total + 1):
            c = ws.cell(title_row, ci, '')
            c.fill   = PatternFill(start_color=NAVY, end_color=NAVY, fill_type='solid')
            c.border = bdr
        if n_total > 1:
            ws.merge_cells(start_row=title_row, end_row=title_row,
                           start_column=1, end_column=n_total)

        month_row  = current_row + 1
        header_row = current_row + 2
        ws.row_dimensions[month_row].height  = 18
        ws.row_dimensions[header_row].height = 17

        for ci in range(1, n_total + 1):
            for ri in (month_row, header_row):
                cell = ws.cell(ri, ci, '')
                cell.fill      = PatternFill(start_color=SLATE, end_color=SLATE, fill_type='solid')
                cell.font      = Font(name=fn, color='FFFFFF', bold=True, size=10)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border    = bdr

        for ci, col_name in enumerate(all_cols, start=1):
            is_idx = (ci <= n_idx)
            is_gt  = (ci == n_total)
            if is_idx or is_gt:
                ws.cell(month_row, ci, col_name)
                ws.merge_cells(start_row=month_row, end_row=header_row,
                               start_column=ci, end_column=ci)

        for ci in range(n_idx + 1, n_total):
            d = all_cols[ci - 1]
            ws.cell(header_row, ci, f"{d.day:02d}")

        month_groups = []
        cur_month, grp_start = None, None
        for ci in range(n_idx + 1, n_total):
            d = all_cols[ci - 1]
            m_val = (d.year, d.month)
            if m_val != cur_month:
                if cur_month is not None:
                    month_groups.append((cur_month, grp_start, ci - 1))
                cur_month, grp_start = m_val, ci
        if cur_month is not None:
            month_groups.append((cur_month, grp_start, n_total - 1))

        for (yr, mo), start_c, end_c in month_groups:
            ws.cell(month_row, start_c, calendar.month_name[mo])
            if end_c > start_c:
                ws.merge_cells(start_row=month_row, end_row=month_row,
                               start_column=start_c, end_column=end_c)

        order_created_map = {}
        if 'ORDER ID' in df.columns and 'CREATED DATE' in df.columns:
            parsed_created = pd.to_datetime(df['CREATED DATE'], dayfirst=True, format='mixed', errors='coerce')
            for order_id, dt in zip(df['ORDER ID'].astype(str).str.strip(), parsed_created):
                if pd.notna(dt):
                    order_created_map[order_id] = dt.date()

        day_totals  = {d: 0 for d in active_days}
        grand_total = 0
        data_start  = current_row + 3

        for ri, row in agg.iterrows():
            r = data_start + ri
            ws.row_dimensions[r].height = 15
            gt_val = int(row.get('Grand Total', 0))
            grand_total += gt_val

            is_overdue = False
            is_overdue_7days = False
            status_code = None
            if 'ORDER ID' in row:
                oid = str(row['ORDER ID']).strip()
                if oid in order_created_map:
                    age_days = (datetime.now().date() - order_created_map[oid]).days
                    if age_days > 1:
                        is_overdue = True
                    if age_days > 7:
                        is_overdue_7days = True
                status_code = order_status_map.get(oid)

            for ci, col in enumerate(all_cols, start=1):
                val  = row.get(col, '')
                cell = ws.cell(r, ci, val if val != '' else None)
                cell.border = bdr
                
                cell_fill = None
                if status_code in ('420', '472') and is_overdue_7days:
                    cell_fill = PatternFill(start_color='FFEBEB', end_color='FFEBEB', fill_type='solid')
                    cell.font = Font(name=fn, size=10, bold=True)
                elif status_code in ('420', '472'):
                    cell_fill = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid')
                    cell.font = Font(name=fn, size=10, bold=True)
                elif is_overdue:
                    cell_fill = PatternFill(start_color='FFEBEB', end_color='FFEBEB', fill_type='solid')
                    cell.font = Font(name=fn, size=10, bold=True)
                elif status_code in ('500', '520', '540'):
                    cell_fill = PatternFill(start_color='FCE4D6', end_color='FCE4D6', fill_type='solid')
                    cell.font = Font(name=fn, size=10, bold=True)
                else:
                    cell.font = Font(name=fn, size=10)

                if cell_fill:
                    cell.fill = cell_fill

                if col in active_days:
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    if isinstance(val, (int, float)) and val:
                        day_totals[col] = day_totals.get(col, 0) + int(val)
                elif col == 'Grand Total':
                    cell.font      = Font(name=fn, color=RED, bold=True, size=10)
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                else:
                    cell.alignment = Alignment(horizontal='left', vertical='center')

        gt_row = data_start + len(agg)
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

        for ci, col in enumerate(all_cols, start=1):
            letter = get_column_letter(ci)
            if col == 'Grand Total':
                ws.column_dimensions[letter].width = 20
            elif col == 'ZONE':
                ws.column_dimensions[letter].width = 9
            elif isinstance(col, __import__('datetime').date):
                ws.column_dimensions[letter].width = 7
            elif col in ('Cus name', 'RECEIVER'):
                max_len = max(
                    (len(str(ws.cell(r_iter, ci).value or ''))
                     for r_iter in range(current_row, gt_row + 1)), default=20)
                existing = ws.column_dimensions[letter].width or 0
                ws.column_dimensions[letter].width = max(existing, min(max(max_len + 3, 22), 50))
            elif col == 'Phone':
                max_len = max(
                    (len(str(ws.cell(r_iter, ci).value or ''))
                     for r_iter in range(current_row, gt_row + 1)), default=14)
                existing = ws.column_dimensions[letter].width or 0
                ws.column_dimensions[letter].width = max(existing, min(max(max_len + 3, 16), 35))
            else:
                max_len = max(
                    (len(str(ws.cell(r_iter, ci).value or ''))
                     for r_iter in range(current_row, gt_row + 1)), default=8)
                existing = ws.column_dimensions[letter].width or 0
                ws.column_dimensions[letter].width = max(existing, min(max(max_len + 2, 10), 28))

        current_row = gt_row + 3

    wb.save(out_path)
    return out_path
