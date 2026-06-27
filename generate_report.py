import pandas as pd
import os
import json
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

# ── Status codes ───────────────────────────────────────────────────────────────
CLASSIFY_LABEL = {'Pickup': 'Pickup', 'Delivery': 'Delivery', 'Pending': 'Pending'}

# Max index cols across all report types (for stacked/long alignment)
MAX_INDEX = 6  # Pickup has most: ZONE, HANDLE, CURRENT PO, ORDER ID, Cus name, Phone

REPORT_COLS = {
    'Pickup':   ['ZONE', 'POST OFFICE HANDLE', 'CURRENT POST OFFICE', 'ORDER ID', 'Cus name', 'Phone'],
    'Delivery': ['ZONE', 'POST OFFICE HANDLE', 'CURRENT POST OFFICE', 'ORDER ID', 'RECEIVER'],
    'Pending':  ['ZONE', 'CURRENT POST OFFICE', 'ORDER ID', 'REMARK'],
}

REPORT_FILTER_COLS = {
    'Pickup':   'POST OFFICE HANDLE',
    'Delivery': 'POST OFFICE HANDLE',
    'Pending':  'POST OFFICE HANDLE',
}

GAP_COLS = 1  # gap between side-by-side tables in wide mode

# Fixed column widths (Excel units)
W_DAY   = 4.5   # day columns e.g. "01","02"
W_ZONE  = 9.0   # ZONE
W_GT    = 18.0  # Grand Total
W_MIN   = 8.0
W_MAX   = 38.0

# Highlight threshold: rows older than this many days get highlighted
HIGHLIGHT_OVER_DAYS = 1
HIGHLIGHT_COLOR = 'FFEBEB'

# Pending remark: classify WHY it's pending based on status code
PENDING_REMARK_MAP = {
    '306': 'Assign deliver',
    '309': 'Assign deliver',
    '311': 'Assign deliver',
    '210': 'Handover Mega truck',
    '300': 'Handover Mega truck',
    '302': 'Handover Mega truck',
    '310': 'Handover Mega truck',
    '500': 'Handover Mega truck',
}


# ── Helpers ────────────────────────────────────────────────────────────────────

def normalize_code(v):
    return '' if pd.isna(v) else str(v).strip().upper()

def clean_text(v):
    if pd.isna(v):
        return ''
    s = str(v).strip()
    return '' if s.lower() == '(blank)' else s

def normalize_id(v):
    if pd.isna(v):
        return ''
    s = str(v).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s

def reference_column(df, keys):
    return next((c for c in df.columns if any(k in c for k in keys)), None)

def load_test_order_ids(cfg):
    test_ids = set()
    
    # Try loading from text file
    txt_path = "test_bills.txt"
    if os.path.exists(txt_path):
        try:
            with open(txt_path, "r", encoding="utf-8") as f:
                for line in f:
                    val = line.strip()
                    if val:
                        test_ids.add(val)
        except Exception as e:
            print(f"Error reading test_bills.txt: {e}")

    # Try loading from delayed bills JSON
    delay_path = "delayed_bills.json"
    if os.path.exists(delay_path):
        try:
            with open(delay_path, "r", encoding="utf-8") as f:
                delayed = json.load(f)
            
            today = datetime.now().date()
            updated_delayed = {}
            for bill_id, exp_date_str in delayed.items():
                try:
                    exp_date = datetime.strptime(exp_date_str, "%Y-%m-%d").date()
                    if today < exp_date:
                        test_ids.add(str(bill_id))
                        updated_delayed[bill_id] = exp_date_str
                except Exception:
                    pass
            
            # Prune expired bills
            if len(updated_delayed) != len(delayed):
                with open(delay_path, "w", encoding="utf-8") as f:
                    json.dump(updated_delayed, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"Error reading delayed_bills.json: {e}")

    tr_cfg = cfg.get("test_receipts", {})
    if not tr_cfg.get("enabled"):
        return test_ids

    path = tr_cfg.get("path")
    if not path or not os.path.exists(path):
        return test_ids

    if path.lower().endswith((".xlsx", ".xls")):
        df_test = pd.read_excel(path, dtype=str)
    else:
        df_test = pd.read_csv(path, dtype=str, keep_default_na=False)
    df_test = df_test.fillna("")

    order_col = next(
        (
            c for c in df_test.columns
            if "order" in str(c).lower()
            or "phi" in str(c).lower()
            or "shipment" in str(c).lower()
        ),
        df_test.columns[0] if len(df_test.columns) else None,
    )
    if not order_col:
        return test_ids

    excel_ids = {
        str(v).strip()
        for v in df_test[order_col].tolist()
        if str(v).strip() and str(v).strip().lower() != "nan"
    }
    
    test_ids.update(excel_ids)
    return test_ids

def build_simple_reference(df_ref):
    key_col      = reference_column(df_ref, ['current_post_office', 'dealer', 'code'])
    handle_col   = reference_column(df_ref, ['post_office_handle', 'responsible', 'handle'])
    customer_col = reference_column(df_ref, ['customer_name', 'agent', 'showroom'])
    phone_col    = reference_column(df_ref, ['phone'])
    if not key_col:
        raise ValueError("Reference file missing current_post_office column.")
    ref = pd.DataFrame({
        'current_post_office': df_ref[key_col].apply(normalize_code),
        'post_office_handle':  df_ref[handle_col].apply(normalize_code) if handle_col else '',
        'customer_name':       df_ref[customer_col].apply(clean_text)   if customer_col else '',
        'phone':               df_ref[phone_col].apply(clean_text)       if phone_col else '',
    })
    ref = ref[ref['current_post_office'] != ''].copy()
    ref['_s'] = ref.apply(lambda r: sum(1 for v in r if clean_text(v)), axis=1)
    ref = ref.sort_values('_s', ascending=False).drop_duplicates('current_post_office', keep='first').drop(columns=['_s'])
    ref['_ref_post_office_key'] = ref['current_post_office']
    return ref

def get_zone(po, zone_mapping):
    po = str(po).strip().upper()
    if po in zone_mapping.get('by_post_office', {}):
        return zone_mapping['by_post_office'][po]
    for prefix, zone in zone_mapping.get('by_prefix', {}).items():
        if po.startswith(prefix.upper()):
            return zone
    return zone_mapping.get('default_zone', 'Zone?')

def load_reference(ref_path):
    if ref_path.lower().endswith('.csv'):
        df = pd.read_csv(ref_path, dtype=str, keep_default_na=False)
    else:
        df = pd.read_excel(ref_path, dtype=str)
    df.columns = [str(c).strip().lower() for c in df.columns]
    df = df.dropna(how='all').fillna('').replace(r'^\s*\(blank\)\s*$', '', regex=True)
    return build_simple_reference(df)


# ── Row builder ────────────────────────────────────────────────────────────────

def build_section_rows(df_h, index_cols, day_cols, date_col):
    """
    Returns (rows, total, active_day_cols).
    active_day_cols = only days that have >=1 order in this section.
    """
    df = df_h.copy()
    for col in index_cols:
        if col not in df.columns:
            df[col] = ''

    if date_col and date_col in df.columns:
        parsed = pd.to_datetime(df[date_col], dayfirst=True, format='mixed', errors='coerce')
        df['_day'] = parsed.dt.strftime('%d').fillna('')
    else:
        df['_day'] = ''

    days_present = set(df['_day'].unique()) - {'', 'NaT', 'nan'}
    active_days = [d for d in day_cols if d in days_present]

    if not active_days:
        footer = {col: '' for col in index_cols}
        footer[index_cols[0]] = 'Grand Total'
        footer['Grand Total'] = 0
        return [footer], 0, []

    for d in active_days:
        df[d] = (df['_day'] == d).astype(int)
    df['Grand Total'] = 1

    agg = df.groupby(index_cols, sort=False, dropna=False)[active_days + ['Grand Total']].sum().reset_index()
    for d in active_days:
        agg[d] = agg[d].apply(lambda v: int(v) if v > 0 else '')

    total = int(agg['Grand Total'].sum())

    footer = {col: '' for col in index_cols}
    footer[index_cols[0]] = 'Grand Total'
    for d in active_days:
        footer[d] = int(sum(v for v in agg[d] if isinstance(v, (int, float)) and v != ''))
    footer['Grand Total'] = total

    rows = agg.to_dict('records')
    rows.append(footer)
    return rows, total, active_days


# ── Style helpers ──────────────────────────────────────────────────────────────

def _fill(h):
    return PatternFill(start_color=h, end_color=h, fill_type='solid')

def _font(name, color='000000', bold=False, size=10):
    return Font(name=name, color=color, bold=bold, size=size)

def _align(h='center', v='center', wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def _border():
    t = Side(style='thin', color='BFBFBF')
    return Border(left=t, right=t, top=t, bottom=t)


# ── Table writer ───────────────────────────────────────────────────────────────

def _write_table(ws, start_row, start_col, report_name, rows, index_cols, active_days, dc, handle="", show_top_title=False, max_index=0, order_created_map=None):
    """
    Write one report section at (start_row, start_col).
    Returns (next_free_row, next_free_col_after_block).
    """
    fn     = dc.get('font_name',         'Aptos Narrow')
    t_fg   = dc.get('title_color',       'FFFFFF')
    t_bg   = dc.get('title_fill_color',  '0F172A')
    h_bg   = dc.get('header_fill_color', '1E293B')
    h_fg   = dc.get('header_font_color', 'FFFFFF')
    z_bg   = dc.get('zone_fill_color',   'FFFFFF')
    z_fg   = dc.get('zone_font_color',   '0F172A')
    g_bg   = dc.get('group_fill_color',  'F8FAFC')
    g_fg   = dc.get('group_font_color',  '0F172A')
    tot_bg = dc.get('total_fill_color',  'F1F5F9')
    RED    = 'EF4444'
    bdr    = _border()

    # Pad index_cols so day columns align perfectly across tables
    padded_index = list(index_cols)
    while len(padded_index) < max_index:
        padded_index.append('')

    all_cols = padded_index + active_days + ['Grand Total']
    n        = len(all_cols)
    end_col  = start_col + n - 1

    # Row 1 — Title (Clean design: Dark background, White text)
    r = start_row
    ws.row_dimensions[r].height = 22
    tc = ws.cell(r, start_col, f"{report_name.upper()} BILL")
    tc.font      = _font(fn, t_fg, bold=True, size=11)  # White title
    tc.fill      = _fill(t_bg)                          # Dark background
    tc.alignment = _align('left')
    tc.border    = bdr
    
    # Pre-style and empty the companion cells BEFORE merging so openpyxl doesn't raise a read-only error
    if n > 1:
        comp_cell = ws.cell(r, start_col + 1, "")
        comp_cell.fill = _fill(t_bg)
        comp_cell.border = bdr
        
    for ci in range(2, n):
        c = ws.cell(r, start_col + ci, "")
        c.fill = _fill(t_bg)
        c.border = bdr

    # Now merge the first two cells safely
    if n > 1:
        ws.merge_cells(start_row=r, end_row=r, start_column=start_col, end_column=start_col + 1)

    # Add red noted title and date to the right of the merged "PICKUP BILL" on top of the title bar
    if n > 3 and show_top_title:
        now_str = datetime.now().strftime('%d %B %Y')
        rc = ws.cell(r, start_col + 2, f"DUE REPORT: {handle.upper()}" if handle else "DUE REPORT")
        rc.font = _font(fn, 'FF8A8A', bold=True, size=11)
        rc.fill = _fill(t_bg)
        rc.alignment = _align('left')
        
        dc_cell = ws.cell(r, start_col + 3, f"DATE: {now_str}")
        dc_cell.font = _font(fn, 'FF8A8A', bold=True, size=11)
        dc_cell.fill = _fill(t_bg)
        dc_cell.alignment = _align('left')

    r += 1

    # Row 2 — Headers
    ws.row_dimensions[r].height = 22
    for ci, col_name in enumerate(all_cols):
        c = ws.cell(r, start_col + ci, col_name)
        c.font      = _font(fn, h_fg, bold=True)
        c.fill      = _fill(h_bg)
        c.alignment = _align('center')
        c.border    = bdr
    r += 1

    # Data rows
    today = datetime.now().date()
    highlight_fill = _fill(HIGHLIGHT_COLOR)

    for row_dict in rows:
        ws.row_dimensions[r].height = 20
        is_total = str(row_dict.get(index_cols[0], '')).strip() == 'Grand Total'

        # Check if this row's order is over 2 days old
        is_overdue = False
        if not is_total and order_created_map:
            order_id = normalize_id(row_dict.get('ORDER ID', ''))
            created_date = order_created_map.get(order_id)
            if created_date and (today - created_date).days > HIGHLIGHT_OVER_DAYS:
                is_overdue = True

        for ci, col_name in enumerate(all_cols):
            val  = row_dict.get(col_name, '')
            cell = ws.cell(r, start_col + ci, val if val != '' else None)
            cell.border = bdr

            if is_total:
                cell.fill      = _fill(tot_bg)
                cell.font      = _font(fn, RED, bold=True)
                cell.alignment = _align('center')
            elif is_overdue:
                # Highlight entire row with light red for overdue orders
                cell.fill      = highlight_fill
                cell.font      = _font(fn, '000000', bold=True)
                cell.alignment = _align('center')
            elif col_name == index_cols[0]:
                cell.fill      = _fill(z_bg)
                cell.font      = _font(fn, z_fg, bold=True)
                cell.alignment = _align('center')
            elif len(index_cols) > 1 and col_name == index_cols[1]:
                cell.fill      = _fill(g_bg)
                cell.font      = _font(fn, g_fg, bold=True)
                cell.alignment = _align('center')
            elif col_name in index_cols:
                # All other index columns (CURRENT POST OFFICE, ORDER ID, etc.) → center
                cell.font      = _font(fn, '1E293B', bold=True)
                cell.alignment = _align('center')
            elif col_name in active_days:
                cell.font      = _font(fn, '1E293B', bold=True)
                cell.alignment = _align('center')
            elif col_name == 'Grand Total':
                cell.font      = _font(fn, RED, bold=True)
                cell.alignment = _align('center')
            else:
                cell.font      = _font(fn, '1E293B', bold=True)
                cell.alignment = _align('center')
        r += 1

    return r, end_col + 1  # (next_free_row, next_free_col)


# ── Column width setter ────────────────────────────────────────────────────────

def _set_col_widths(ws):
    """
    Set fixed widths:
    - Day cols (2-digit header 01-31) → W_DAY (all same, uniform)
    - ZONE col → W_ZONE
    - Grand Total col → W_GT
    - Empty cols → 1.0 (hidden/minimal)
    - Other text cols → auto-fit capped at W_MAX
    """
    # Scan all header rows to classify columns
    day_cols_ci   = set()
    zone_cols_ci  = set()
    gt_cols_ci    = set()
    empty_cols_ci = set()

    for r in range(1, min(ws.max_row + 1, ws.max_row + 1)):
        for c in range(1, ws.max_column + 1):
            try:
                val = str(ws.cell(r, c).value or '').strip()
            except Exception:
                val = ''
            if val.isdigit() and len(val) == 2 and 1 <= int(val) <= 31:
                day_cols_ci.add(c)
            elif val == 'ZONE':
                zone_cols_ci.add(c)
            elif val == 'Grand Total':
                gt_cols_ci.add(c)

    # Detect fully empty columns
    for c in range(1, ws.max_column + 1):
        if all(
            str(ws.cell(r, c).value or '').strip() == ''
            for r in range(1, ws.max_row + 1)
        ):
            empty_cols_ci.add(c)

    for c in range(1, ws.max_column + 1):
        letter = get_column_letter(c)
        if c in day_cols_ci:
            ws.column_dimensions[letter].width = W_DAY
        elif c in zone_cols_ci:
            ws.column_dimensions[letter].width = W_ZONE
        elif c in gt_cols_ci:
            ws.column_dimensions[letter].width = W_GT
        elif c in empty_cols_ci:
            ws.column_dimensions[letter].width = 1.0
        else:
            # Determine header value to apply specific rules
            header_val = ''
            for r_scan in range(1, min(10, ws.max_row + 1)):
                cell_v = str(ws.cell(r_scan, c).value or '').strip()
                if cell_v in ('Cus name', 'RECEIVER', 'Phone', 'ORDER ID', 'CURRENT POST OFFICE', 'POST OFFICE HANDLE', 'REMARK'):
                    header_val = cell_v
                    break
            
            max_len = 0
            for r in range(1, ws.max_row + 1):
                try:
                    v = ws.cell(r, c).value
                    if v:
                        max_len = max(max_len, len(str(v)))
                except Exception:
                    pass
            
            if header_val in ('Cus name', 'RECEIVER'):
                ws.column_dimensions[letter].width = min(max(max_len + 4, 25), 55)
            elif header_val == 'Phone':
                ws.column_dimensions[letter].width = min(max(max_len + 4, 19), 38)
            elif header_val == 'ORDER ID':
                ws.column_dimensions[letter].width = min(max(max_len + 3, 16), 24)
            else:
                ws.column_dimensions[letter].width = min(max(max_len + 2, W_MIN), W_MAX)


# ── Excel builders ─────────────────────────────────────────────────────────────

def build_handle_excel(handle, sections, day_cols, dc, out_path, mode='wide', order_created_map=None):
    """
    sections: list of (report_name, rows, total, index_cols, active_days)
    mode: 'wide' = side by side, 'long' = stacked
    """
    wb = Workbook()
    ws = wb.active
    ws.title = handle[:31]
    fn  = dc.get('font_name', 'Aptos Narrow')

    # Shared active days = union of all section active days, in sorted order
    shared_days = sorted(set(d for _, _, _, _, ad in sections for d in ad))

    # Compute total cols for merge using the widest index cols set
    max_index = max((len(ic) for _, _, _, ic, _ in sections), default=3)
    total_cols = max_index + len(shared_days) + 1  # +1 for Grand Total

    if mode == 'wide':
        cur_col = 1
        for report_name, rows, total, index_cols, active_days in sections:
            _, next_col = _write_table(ws, 1, cur_col, report_name,
                                       rows, index_cols, shared_days, dc, handle=handle, max_index=max_index, order_created_map=order_created_map)
            cur_col = next_col + GAP_COLS
    else:
        # Long mode: Pickup → Delivery → Pending stacked
        r = 1
        for i, (report_name, rows, total, index_cols, active_days) in enumerate(sections):
            next_row, _ = _write_table(ws, r, 1, report_name,
                                       rows, index_cols, shared_days, dc, handle=handle, show_top_title=(i==0), max_index=max_index, order_created_map=order_created_map)
            r = next_row + 1  # 1 blank row gap

    _set_col_widths(ws)
    wb.save(out_path)


def build_final_excel(all_handle_sections, day_cols, dc, out_path, mode='wide', order_created_map=None):
    wb = Workbook()
    ws = wb.active
    ws.title = f"Report {datetime.now().strftime('%d.%m')}"

    # Global shared days across ALL handles
    shared_days = sorted(set(
        d for _, sections in all_handle_sections
        for _, _, _, _, ad in sections
        for d in ad
    ))

    # Compute total cols for merge using the widest index cols set across all sections
    max_index = 3
    for _, sections in all_handle_sections:
        for _, _, _, ic, _ in sections:
            max_index = max(max_index, len(ic))

    cur_row = 1
    for handle, sections in all_handle_sections:
        if mode == 'wide':
            max_rows = max((len(rows) for _, rows, _, _, _ in sections), default=0)
            cur_col  = 1
            for report_name, rows, total, index_cols, active_days in sections:
                _, next_col = _write_table(ws, cur_row, cur_col, report_name,
                                           rows, index_cols, shared_days, dc, handle=handle, max_index=max_index, order_created_map=order_created_map)
                cur_col = next_col + GAP_COLS
            cur_row = cur_row + 2 + max_rows + 2
        else:
            for i, (report_name, rows, total, index_cols, active_days) in enumerate(sections):
                next_row, _ = _write_table(ws, cur_row, 1, report_name,
                                           rows, index_cols, shared_days, dc, handle=handle, show_top_title=(i==0), max_index=max_index, order_created_map=order_created_map)
                cur_row = next_row + 1
            cur_row += 1

    _set_col_widths(ws)
    wb.save(out_path)


# ── Main ───────────────────────────────────────────────────────────────────────

def generate_reports_from_data(export_path, ref_path, output_dir,
                                return_metadata=False, mode='wide', target_handles=None):
    config_path = os.path.join(os.path.dirname(__file__), 'config.json')
    with open(config_path, encoding='utf-8') as f:
        cfg = json.load(f)
    zone_mapping = cfg.get('zone_mapping', {})
    excel_design = cfg.get('excel_design', {})


    df_ref = load_reference(ref_path)

    # Auto-detect sheet
    xl = pd.ExcelFile(export_path)
    sheet = xl.sheet_names[0]
    if len(xl.sheet_names) > 1:
        for sname in xl.sheet_names:
            try:
                s = xl.parse(sname, nrows=3)
                if 'ORDER ID' in s.columns and (
                    'CURRENT STATUS' in s.columns or
                    'is tồn kết nối' in s.columns
                ):
                    sheet = sname
                    break
            except Exception:
                continue
    df = xl.parse(sheet)

    # Date column for day-of-month grouping
    date_col = (
        'CREATED DATE' if 'CREATED DATE' in df.columns else
        'CURRENT TIME' if 'CURRENT TIME' in df.columns else None
    )

    order_created_map = {}
    if date_col and 'ORDER ID' in df.columns:
        parsed_created = pd.to_datetime(df[date_col], dayfirst=True, format='mixed', errors='coerce')
        for order_id, dt in zip(df['ORDER ID'], parsed_created):
            if pd.notna(dt):
                order_created_map[normalize_id(order_id)] = dt.date()

    today = datetime.now().date()

    # Build day_cols = only days that have data in the whole export, + today
    if date_col:
        parsed = pd.to_datetime(df[date_col], dayfirst=True, format='mixed', errors='coerce')
        days_with_data = set(parsed.dropna().dt.strftime('%d').tolist())
    else:
        days_with_data = set()

    today_str = today.strftime('%d')
    # All days 01 → today, keep only those with data OR today
    day_cols = [
        f"{d:02d}" for d in range(1, today.day + 1)
        if f"{d:02d}" in days_with_data or f"{d:02d}" == today_str
    ]

    # Exclude test orders
    test_col = next(
        (c for c in df.columns if str(c).strip().lower() in ('is test', 'đơn test', 'don test')),
        None
    )
    if test_col:
        df = df[
            df[test_col].isna() |
            df[test_col].astype(str).str.strip().isin(['', 'nan', 'NaN', '#N/A'])
        ].copy()

    test_order_ids = load_test_order_ids(cfg)
    if test_order_ids and 'ORDER ID' in df.columns:
        df = df[~df['ORDER ID'].astype(str).str.strip().isin(test_order_ids)].copy()

    # Exclude test orders based on keywords in name/note columns
    if cfg.get("pivot", {}).get("exclude_test", False):
        keywords = cfg["pivot"].get("test_keywords", ["test"])
        if keywords:
            check_cols = [c for c in df.columns if any(k in str(c).lower() for k in ('name', 'note', 'customer', 'sender', 'receiver', 'remark', 'address'))]
            for kw in keywords:
                kw_lower = str(kw).lower().strip()
                for col in check_cols:
                    # Drop rows where the keyword is found in the column
                    df = df[~df[col].astype(str).str.lower().str.contains(kw_lower, na=False)].copy()

    if df.empty:
        raise ValueError("No data left after test/exclusion filters.")

    # Classification column
    classify_col = next(
        (c for c in df.columns if 'tồn kết nối' in str(c).lower()), None
    )

    # Status code
    if 'CURRENT STATUS' in df.columns:
        df['CURRENT STATUS'] = df['CURRENT STATUS'].astype(str).str.strip()
        df['STATUS_CODE'] = df['CURRENT STATUS'].str.extract(r'^(\d{3})')[0]
    else:
        df['STATUS_CODE'] = ''

    if 'CURRENT POST OFFICE' in df.columns:
        df['_po_key'] = df['CURRENT POST OFFICE'].apply(normalize_code)
    else:
        df['_po_key'] = ''

    # POST OFFICE HANDLE + ZONE
    if 'POST OFFICE HANDLE' in df.columns:
        dm = df.copy()
        dm['POST OFFICE HANDLE'] = dm['POST OFFICE HANDLE'].apply(normalize_code)
        if 'CURRENT POST OFFICE' in dm.columns:
            dm['CURRENT POST OFFICE'] = dm['CURRENT POST OFFICE'].apply(normalize_code)
    else:
        dm = pd.merge(df, df_ref, left_on='_po_key', right_on='_ref_post_office_key', how='left')
        if 'CURRENT POST OFFICE' in dm.columns:
            dm['CURRENT POST OFFICE'] = dm['CURRENT POST OFFICE'].apply(normalize_code)
        dm['POST OFFICE HANDLE'] = dm['post_office_handle'].apply(normalize_code)
        if 'CURRENT POST OFFICE' in dm.columns:
            dm.loc[dm['POST OFFICE HANDLE'] == '', 'POST OFFICE HANDLE'] = dm['CURRENT POST OFFICE']

    if 'ZONE' not in dm.columns or dm['ZONE'].isna().all():
        dm['ZONE'] = dm['POST OFFICE HANDLE'].apply(lambda x: get_zone(x, zone_mapping))
    else:
        dm['ZONE'] = dm['ZONE'].apply(clean_text)

    if 'ORDER ID' in dm.columns:
        dm['ORDER ID'] = dm['ORDER ID'].astype(str).str.strip()

    for col in ("RECEIVE POST OFFICE", "DELIVERY POST OFFICE", "CURRENT POST OFFICE"):
        if col not in dm.columns:
            dm[col] = ""
        dm[col] = dm[col].apply(normalize_code)

    # RECEIVER
    if 'RECEIVER' not in dm.columns:
        if 'RECEIVER' in df.columns:
            dm['RECEIVER'] = df['RECEIVER'].apply(
                lambda v: str(v).split(' - ', 1)[1].strip() if ' - ' in str(v) else clean_text(v))
        else:
            dm['RECEIVER'] = ''

    # Cus name / Phone
    cus_src = next((c for c in dm.columns if c.strip() == 'Cus name'), None)
    if cus_src:
        dm['Cus name'] = dm[cus_src].apply(clean_text)
    elif 'SENDER' in dm.columns:
        dm['Cus name'] = dm['SENDER'].apply(
            lambda v: str(v).split(' - ', 1)[1].strip() if ' - ' in str(v) else clean_text(v))
    else:
        dm['Cus name'] = ''

    if 'Phone' not in dm.columns:
        if 'SENDER' in dm.columns:
            dm['Phone'] = dm['SENDER'].apply(
                lambda v: str(v).split(' - ', 1)[0].strip() if ' - ' in str(v) else '')
        else:
            dm['Phone'] = ''

    os.makedirs(output_dir, exist_ok=True)

    # Build status map from config.json
    status_map = {}
    for r in cfg.get("reports", []):
        label = 'Pickup' if 'pickup' in r.get('key', '').lower() else \
                'Delivery' if 'delivery' in r.get('key', '').lower() else \
                'Pending' if 'pending' in r.get('key', '').lower() else r.get('is_label', '')
        for sc in r.get("status_codes", []):
            status_map[str(sc).strip()] = label

    # Globally drop completed statuses from all reports (Pickup, Delivery, Pending)
    if 'STATUS_CODE' in dm.columns:
        dm = dm[~dm['STATUS_CODE'].isin(['410', '201', '520'])].copy()

    # Filter per report type
    type_data = {}
    if classify_col:
        dm["_report_class"] = dm[classify_col].astype(str).str.strip()
    else:
        dm["_report_class"] = dm['STATUS_CODE'].map(status_map).fillna("Unknown")

    if target_handles:
        target_handles = [h.upper() for h in target_handles if h]

    for rn in ['Pickup', 'Delivery', 'Pending']:
        df_t = dm[dm["_report_class"] == CLASSIFY_LABEL[rn]].copy()
        if target_handles:
            filter_col = REPORT_FILTER_COLS[rn]
            if filter_col in df_t.columns:
                df_t = df_t[df_t[filter_col].isin(target_handles)]
        # Add REMARK column for Pending and remove completed statuses
        if rn == 'Pending' and 'STATUS_CODE' in df_t.columns:
            def _pending_remark(row):
                sc = str(row.get('STATUS_CODE', '')).strip()
                return PENDING_REMARK_MAP.get(sc, 'Unknown')
            df_t['REMARK'] = df_t.apply(_pending_remark, axis=1)
        type_data[rn] = df_t

    # Gather all handles
    all_handles = set()
    for rn in ['Pickup', 'Delivery', 'Pending']:
        df_t = type_data[rn]
        filter_col = REPORT_FILTER_COLS[rn]
        if filter_col in df_t.columns:
            all_handles.update(df_t[filter_col].dropna().unique())
            
    unique_handles = sorted(list(h for h in all_handles if str(h).strip()))
    if target_handles:
        for th in target_handles:
            if th not in unique_handles:
                unique_handles.append(th)
        unique_handles = sorted(unique_handles)

    handle_results = []
    all_handle_sections = []
    overall = {'Pickup': 0, 'Delivery': 0, 'Pending': 0}

    for handle in unique_handles:
        sections = []
        counts = {}
        for rn in ['Pickup', 'Delivery', 'Pending']:
            df_t = type_data[rn]
            filter_col = REPORT_FILTER_COLS[rn]
            if filter_col in df_t.columns:
                df_h = df_t[df_t[filter_col] == handle].copy()
            else:
                df_h = pd.DataFrame()

            icols = REPORT_COLS[rn]
            if df_h.empty:
                counts[rn] = 0
                continue
                
            sort_cols = [c for c in ['ZONE', filter_col, 'ORDER ID'] if c in df_h.columns]
            if sort_cols:
                df_h = df_h.sort_values(by=sort_cols)

            rows, total, active_days = build_section_rows(df_h, icols, day_cols, date_col)
            counts[rn] = total
            overall[rn] += total
            if total > 0:
                sections.append((rn, rows, total, icols, active_days))

        total_handle = sum(counts.values())
        if total_handle == 0 and not (target_handles and handle in target_handles):
            continue
            
        if total_handle > 0:
            all_handle_sections.append((handle, sections))
        
        handle_files = []
        for rn, rows, total, icols, active_days in sections:
            tmp_xlsx = os.path.join(output_dir, f"Report_{handle}_{rn}_{today.strftime('%d_%m_%Y_%H%M%S')}.xlsx")
            build_handle_excel(f"Report_{handle}_{rn}", [(rn, rows, total, icols, active_days)], day_cols, excel_design, tmp_xlsx, mode=mode, order_created_map=order_created_map)
            handle_files.append({'path': tmp_xlsx, 'handle': handle})

        remark = (
            f"{handle}  |  "
            + "  |  ".join(f"{t}: {counts.get(t,0)}" for t in ['Pickup','Delivery','Pending'])
            + f"  |  Total: {total_handle}"
        )

        handle_results.append({
            'handle':        handle,
            'handle_counts': counts,
            'handle_files':  handle_files,
            'remark':        remark,
            'sections':      sections,
        })

    final_xlsx = os.path.join(output_dir, f"Report_All_{today.strftime('%d_%m_%Y')}.xlsx")
    if all_handle_sections:
        build_final_excel(all_handle_sections, day_cols, excel_design, final_xlsx, mode=mode, order_created_map=order_created_map)
    else:
        wb = Workbook()
        wb.save(final_xlsx)
        
    if not handle_results:
        handle_str = 'ALL' if not target_handles else ', '.join(target_handles)
        handle_results = [{
            'handle':        handle_str,
            'handle_counts': overall,
            'handle_files':  [{'path': final_xlsx, 'handle': handle_str}],
            'remark':        f"No data found for {handle_str}",
            'sections':      [],
        }]

    grand_total = sum(overall.values())
    summary = "\n".join([
        f"📋 Daily Report  {datetime.now().strftime('%d/%m/%Y %H:%M')}",
        f"Pickup: {overall['Pickup']}  |  Delivery: {overall['Delivery']}  |  Pending: {overall['Pending']}",
        f"Grand Total: {grand_total}",
    ])

    result = {
        'handle_results':  handle_results,
        'final_xlsx':      final_xlsx,
        'summary_caption': summary,
        'overall_counts':  overall,
        'type_data':       type_data,
        'day_cols':        day_cols,
        'cur_time_col':    date_col,
    }
    return result if return_metadata else [final_xlsx]

