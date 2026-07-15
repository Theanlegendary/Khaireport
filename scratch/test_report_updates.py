import os
import json
import pandas as pd
from datetime import datetime
from openpyxl import Workbook
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

# Load current config
with open("c:/Users/DELL/Desktop/daily_push/config.json", encoding="utf-8") as f:
    cfg = json.load(f)

import generate_report
import excel_to_image

# 1. Define the Khmer header map and helper function
HEADER_KHMER_MAP = {
    'ZONE': 'តំបន់',
    'POST OFFICE HANDLE': 'ប៉ុស្តិ៍ទទួលខុសត្រូវ',
    'CURRENT POST OFFICE': 'ប៉ុស្តិ៍បច្ចុប្បន្ន',
    'ORDER ID': 'លេខបុង / លេខកូដបញ្ជាទិញ',
    'Cus name': 'ឈ្មោះអតិថិជន',
    'Phone': 'លេខទូរស័ព្ទ',
    'RECEIVER': 'ឈ្មោះអតិថិជន',
    'REMARK': 'ស្ថានភាពបច្ចុប្បន្ន',
    'NEXT_ACTION': 'សកម្មភាពត្រូវធ្វើ',
    'Grand Total': 'សរុប',
}

def translate_header(col_name):
    return HEADER_KHMER_MAP.get(str(col_name).strip(), col_name)

def get_next_action(status_code):
    sc = str(status_code).strip()
    if sc in ('306', '309', '311'):
        return 'ត្រូវចាត់ចែងអ្នកដឹក'
    elif sc in ('300', '302'):
        return 'ត្រូវស្កេនទទួលអីវ៉ាន់'
    elif sc in ('210', '310', '500'):
        return 'រង់ចាំឡានដឹកមកដល់'
    return ''

# Monkey-patch or override the key structures in generate_report
generate_report.REPORT_COLS = {
    'Pickup':   ['ZONE', 'POST OFFICE HANDLE', 'CURRENT POST OFFICE', 'ORDER ID', 'Cus name', 'Phone'],
    'Delivery': ['ZONE', 'POST OFFICE HANDLE', 'CURRENT POST OFFICE', 'ORDER ID', 'RECEIVER'],
    'កំពុងលើផ្លូវ (Transit)':  ['ZONE', 'CURRENT POST OFFICE', 'ORDER ID', 'REMARK', 'NEXT_ACTION'],
    'ដល់ប៉ុស្តិ៍ (Action Needed)': ['ZONE', 'CURRENT POST OFFICE', 'ORDER ID', 'REMARK', 'NEXT_ACTION'],
    'Pending': ['ZONE', 'CURRENT POST OFFICE', 'ORDER ID', 'REMARK', 'NEXT_ACTION'],
}

generate_report.REPORT_FILTER_COLS = {
    'Pickup':   'POST OFFICE HANDLE',
    'Delivery': 'POST OFFICE HANDLE',
    'Pending':  'POST OFFICE HANDLE',
    'កំពុងលើផ្លូវ (Transit)':  'POST OFFICE HANDLE',
    'ដល់ប៉ុស្តិ៍ (Action Needed)': 'POST OFFICE HANDLE',
}

# Override translate_header in generate_report module
generate_report.translate_header = translate_header
generate_report.get_next_action = get_next_action

# Let's write the modified build_sheet_section function that uses translate_header
def _write_table_khmer(ws, start_row, start_col, report_name, rows, index_cols, day_cols, dc, handle="", show_top_title=True, max_index=6, order_created_map=None):
    fn     = dc.get('font_name', 'Segoe UI')
    h_bg   = dc.get('header_bg', 'E2E8F0')
    h_fg   = dc.get('header_fg', '1E293B')
    z_bg   = dc.get('zone_bg', 'F1F5F9')
    z_fg   = dc.get('zone_fg', '475569')
    g_bg   = dc.get('group_bg', 'F8FAFC')
    g_fg   = dc.get('group_fg', '475569')
    tot_bg = dc.get('total_bg', 'F1F5F9')
    title_bg = dc.get('title_bg', '0F172A')
    
    RED = '991B1B'
    bdr = Border(left=Side(style='thin', color='BFBFBF'), right=Side(style='thin', color='BFBFBF'),
                 top=Side(style='thin', color='BFBFBF'), bottom=Side(style='thin', color='BFBFBF'))

    # Pad index columns to max_index
    padded_index = index_cols + [''] * (max_index - len(index_cols))
    all_cols = padded_index + day_cols + ['Grand Total']
    n = len(all_cols)

    r = start_row
    if show_top_title:
        # Title row
        ws.row_dimensions[r].height = 28
        ws.merge_cells(start_row=r, end_row=r, start_column=start_col, end_column=start_col + n - 1)
        tc = ws.cell(r, start_col, f" {handle} - {report_name.upper()} ".strip())
        tc.fill = PatternFill(start_color=title_bg, end_color=title_bg, fill_type='solid')
        tc.font = Font(name=fn, color='FFFFFF', bold=True, size=11)
        tc.alignment = Alignment(horizontal='left', vertical='center')
        r += 1

    # Header Row
    ws.row_dimensions[r].height = 20
    ws.row_dimensions[r + 1].height = 20

    # 1. Pre-fill and style all cells in both header rows
    for ci in range(n):
        col_idx = start_col + ci
        for row_idx in (r, r + 1):
            cell = ws.cell(row_idx, col_idx)
            cell.fill = PatternFill(start_color=h_bg, end_color=h_bg, fill_type='solid')
            cell.font = Font(name="Khmer OS Battambang", color=h_fg, bold=True, size=9)
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = bdr

    # 2. Vertically merge index columns and Grand Total column across both rows
    for ci, col_name in enumerate(all_cols):
        col_idx = start_col + ci
        is_index = (ci < len(padded_index))
        is_gt = (ci == n - 1)
        if is_index or is_gt:
            ws.cell(r, col_idx, translate_header(col_name))
            ws.merge_cells(start_row=r, end_row=r + 1, start_column=col_idx, end_column=col_idx)

    # 3. Write day numbers
    for ci in range(len(padded_index), n - 1):
        col_idx = start_col + ci
        col_name = all_cols[ci]  # datetime.date
        ws.cell(r + 1, col_idx, f"{col_name.day:02d}")

    # 4. Group day columns by month
    month_groups = []
    current_month = None
    group_start = None
    for ci in range(len(padded_index), n - 1):
        col_name = all_cols[ci]
        m_val = (col_name.year, col_name.month)
        if m_val != current_month:
            if current_month is not None:
                month_groups.append((current_month, group_start, start_col + ci - 1))
            current_month = m_val
            group_start = start_col + ci
    if current_month is not None:
        month_groups.append((current_month, group_start, start_col + n - 2))

    import calendar
    for (yr, mo), start_c, end_c in month_groups:
        ws.cell(r, start_c, calendar.month_name[mo])
        if end_c > start_c:
            ws.merge_cells(start_row=r, end_row=r, start_column=start_c, end_column=end_c)

    r += 2

    # Data rows
    today = datetime.now().date()
    highlight_fill = PatternFill(start_color=generate_report.HIGHLIGHT_COLOR, end_color=generate_report.HIGHLIGHT_COLOR, fill_type='solid')

    for row_dict in rows:
        ws.row_dimensions[r].height = 20
        is_total = str(row_dict.get(index_cols[0], '')).strip() == 'Grand Total'

        is_overdue = False
        if not is_total and order_created_map:
            order_id = generate_report.normalize_id(row_dict.get('ORDER ID', ''))
            created_date = order_created_map.get(order_id)
            if created_date and (today - created_date).days > generate_report.HIGHLIGHT_OVER_DAYS:
                is_overdue = True

        for ci, col_name in enumerate(all_cols):
            val  = row_dict.get(col_name, '')
            cell = ws.cell(r, start_col + ci, val if val != '' else None)
            cell.border = bdr

            if is_total:
                cell.fill      = PatternFill(start_color=tot_bg, end_color=tot_bg, fill_type='solid')
                cell.font      = Font(name=fn, color=RED, bold=True, size=10)
                cell.alignment = Alignment(horizontal='center', vertical='center')
            elif is_overdue:
                cell.fill      = highlight_fill
                cell.font      = Font(name=fn, color='000000', bold=False, size=10)
                cell.alignment = Alignment(horizontal='center', vertical='center')
            elif col_name == index_cols[0]:
                cell.fill      = PatternFill(start_color=z_bg, end_color=z_bg, fill_type='solid')
                cell.font      = Font(name=fn, color=z_fg, bold=False, size=10)
                cell.alignment = Alignment(horizontal='center', vertical='center')
            elif len(index_cols) > 1 and col_name == index_cols[1]:
                cell.fill      = PatternFill(start_color=g_bg, end_color=g_bg, fill_type='solid')
                cell.font      = Font(name=fn, color=g_fg, bold=False, size=10)
                cell.alignment = Alignment(horizontal='center', vertical='center')
            elif col_name in index_cols:
                cell.font      = Font(name="Khmer OS Battambang" if col_name == 'NEXT_ACTION' or col_name == 'REMARK' else fn, color='1E293B', bold=False, size=9)
                cell.alignment = Alignment(horizontal='center', vertical='center')
            elif col_name in day_cols:
                cell.font      = Font(name=fn, color='1E293B', bold=False, size=10)
                cell.alignment = Alignment(horizontal='center', vertical='center')
            elif col_name == 'Grand Total':
                cell.font      = Font(name=fn, color=RED, bold=True, size=10)
                cell.alignment = Alignment(horizontal='center', vertical='center')
            else:
                cell.font      = Font(name=fn, color='1E293B', size=10)
                cell.alignment = Alignment(horizontal='center', vertical='center')

        r += 1

    return r, start_col + n

# Inject the Khmer write_table implementation
generate_report._write_table = _write_table_khmer

# Let's rewrite generate_reports_from_data with the split Pending section logic
def generate_reports_from_data_khmer(export_path, ref_path, output_dir, target_handles=None, mode='wide', return_metadata=False):
    # Load and clean source files
    if export_path.lower().endswith(('.xlsx', '.xls')):
        df = pd.read_excel(export_path, dtype=str)
    else:
        df = pd.read_csv(export_path, dtype=str, keep_default_na=False)

    df = df.fillna('')
    df.columns = [str(c).strip().upper() for c in df.columns]

    df_ref = generate_report.load_reference(ref_path)
    
    with open("c:/Users/DELL/Desktop/daily_push/config.json", encoding="utf-8") as f:
        cfg = json.load(f)

    # Resolve date column
    date_col = generate_report.reference_column(df, ['order_created_date', 'created_at', 'date', 'created'])
    order_created_map = {}
    if date_col:
        for _, row in df.iterrows():
            oid = generate_report.normalize_id(row.get('ORDER ID', ''))
            dt_val = row.get(date_col)
            if oid and dt_val:
                try:
                    order_created_map[oid] = pd.to_datetime(dt_val, dayfirst=True, format='mixed', errors='coerce').date()
                except Exception:
                    pass

    # Zone mapping
    zone_mapping = cfg.get('zone_mapping', {'by_post_office': {}, 'by_prefix': {}, 'default_zone': 'Zone?'})

    # Classify column
    classify_col = generate_report.reference_column(df, ['report_classification', 'classification', 'type'])

    # Build _po_key
    po_col = generate_report.reference_column(df, ['current_post_office', 'po', 'post_office'])
    if po_col:
        df['_po_key'] = df[po_col].apply(generate_report.normalize_code)
    else:
        df['_po_key'] = ''

    # Merge ref
    if 'POST OFFICE HANDLE' in df.columns:
        dm = df.copy()
        dm['POST OFFICE HANDLE'] = dm['POST OFFICE HANDLE'].apply(generate_report.normalize_code)
        if 'CURRENT POST OFFICE' in dm.columns:
            dm['CURRENT POST OFFICE'] = dm['CURRENT POST OFFICE'].apply(generate_report.normalize_code)
    else:
        dm = pd.merge(df, df_ref, left_on='_po_key', right_on='_ref_post_office_key', how='left')
        if 'CURRENT POST OFFICE' in dm.columns:
            dm['CURRENT POST OFFICE'] = dm['CURRENT POST OFFICE'].apply(generate_report.normalize_code)
        dm['POST OFFICE HANDLE'] = dm['post_office_handle'].apply(generate_report.normalize_code)
        if 'CURRENT POST OFFICE' in dm.columns:
            dm.loc[dm['POST OFFICE HANDLE'] == '', 'POST OFFICE HANDLE'] = dm['CURRENT POST OFFICE']

    if 'ZONE' not in dm.columns or dm['ZONE'].isna().all():
        dm['ZONE'] = dm['POST OFFICE HANDLE'].apply(lambda x: generate_report.get_zone(x, zone_mapping))
    else:
        dm['ZONE'] = dm['ZONE'].apply(generate_report.clean_text)

    if 'ORDER ID' in dm.columns:
        dm['ORDER ID'] = dm['ORDER ID'].astype(str).str.strip()

    for col in ("RECEIVE POST OFFICE", "DELIVERY POST OFFICE", "CURRENT POST OFFICE"):
        if col not in dm.columns:
            dm[col] = ""
        dm[col] = dm[col].apply(generate_report.normalize_code)

    if 'RECEIVER' not in dm.columns:
        if 'RECEIVER' in df.columns:
            dm['RECEIVER'] = df['RECEIVER'].apply(
                lambda v: str(v).split(' - ', 1)[1].strip() if ' - ' in str(v) else generate_report.clean_text(v))
        else:
            dm['RECEIVER'] = ''

    cus_src = next((c for c in dm.columns if c.strip() == 'Cus name'), None)
    if cus_src:
        dm['Cus name'] = dm[cus_src].apply(generate_report.clean_text)
    elif 'SENDER' in dm.columns:
        dm['Cus name'] = dm['SENDER'].apply(
            lambda v: str(v).split(' - ', 1)[1].strip() if ' - ' in str(v) else generate_report.clean_text(v))
    else:
        dm['Cus name'] = ''

    if 'Phone' not in dm.columns:
        if 'SENDER' in dm.columns:
            dm['Phone'] = dm['SENDER'].apply(
                lambda v: str(v).split(' - ', 1)[0].strip() if ' - ' in str(v) else '')
        else:
            dm['Phone'] = ''

    os.makedirs(output_dir, exist_ok=True)

    status_map = {}
    for r in cfg.get("reports", []):
        label = 'Pickup' if 'pickup' in r.get('key', '').lower() else \
                'Delivery' if 'delivery' in r.get('key', '').lower() else \
                'Pending' if 'pending' in r.get('key', '').lower() else r.get('is_label', '')
        for sc in r.get("status_codes", []):
            status_map[str(sc).strip()] = label

    if 'STATUS_CODE' in dm.columns:
        dm = dm[~dm['STATUS_CODE'].isin(['410', '201', '520'])].copy()

    type_data = {}
    if classify_col:
        dm["_report_class"] = dm[classify_col].astype(str).str.strip()
    else:
        dm["_report_class"] = dm['STATUS_CODE'].map(status_map).fillna("Unknown")

    if target_handles:
        target_handles = [h.upper() for h in target_handles if h]

    # Map Pending details first
    for rn in ['Pickup', 'Delivery', 'Pending']:
        df_t = dm[dm["_report_class"] == generate_report.CLASSIFY_LABEL[rn]].copy()
        if target_handles:
            filter_col = generate_report.REPORT_FILTER_COLS[rn]
            if filter_col in df_t.columns:
                df_t = df_t[df_t[filter_col].isin(target_handles)]
        if rn == 'Pending' and 'STATUS_CODE' in df_t.columns:
            def _pending_remark(row):
                sc = str(row.get('STATUS_CODE', '')).strip()
                return generate_report.PENDING_REMARK_MAP.get(sc, 'Unknown')
            def _pending_next_action(row):
                sc = str(row.get('STATUS_CODE', '')).strip()
                return get_next_action(sc)
            df_t['REMARK'] = df_t.apply(_pending_remark, axis=1)
            df_t['NEXT_ACTION'] = df_t.apply(_pending_next_action, axis=1)
        type_data[rn] = df_t

    all_handles = set()
    for rn in ['Pickup', 'Delivery', 'Pending']:
        df_t = type_data[rn]
        filter_col = generate_report.REPORT_FILTER_COLS[rn]
        if filter_col in df_t.columns:
            all_handles.update(df_t[filter_col].dropna().unique())
            
    unique_handles = sorted(list(h for h in all_handles if str(h).strip()))
    if target_handles:
        for th in target_handles:
            if th not in unique_handles:
                unique_handles.append(th)
        unique_handles = sorted(unique_handles)

    # 14 days columns
    day_cols = []
    for di in range(cfg["api"]["date_range_days"]):
        day_cols.append(datetime.now().date() - timedelta(days=di))
    day_cols = sorted(day_cols)

    handle_results = []
    all_handle_sections = []
    overall = {'Pickup': 0, 'Delivery': 0, 'Pending': 0}

    excel_design = cfg.get("excel_design", {})

    for handle in unique_handles:
        sections = []
        counts = {'Pickup': 0, 'Delivery': 0, 'Pending': 0}
        for rn in ['Pickup', 'Delivery', 'Pending']:
            df_t = type_data[rn]
            filter_col = generate_report.REPORT_FILTER_COLS[rn]
            if filter_col in df_t.columns:
                df_h = df_t[df_t[filter_col] == handle].copy()
            else:
                df_h = pd.DataFrame()

            if df_h.empty:
                continue

            sort_cols = [c for c in ['ZONE', filter_col, 'ORDER ID'] if c in df_h.columns]
            if sort_cols:
                df_h = df_h.sort_values(by=sort_cols)

            if rn == 'Pending':
                transit_codes = {'210', '310', '500'}
                df_h_transit = df_h[df_h['STATUS_CODE'].astype(str).str.strip().isin(transit_codes)].copy()
                df_h_action = df_h[~df_h['STATUS_CODE'].astype(str).str.strip().isin(transit_codes)].copy()

                icols = generate_report.REPORT_COLS['Pending']
                
                # Transit section
                if not df_h_transit.empty:
                    rows, total, active_days = generate_report.build_section_rows(df_h_transit, icols, day_cols, date_col)
                    sections.append(('កំពុងលើផ្លូវ (Transit)', rows, total, icols, active_days))
                    counts['Pending'] += total
                    overall['Pending'] += total
                    
                # Action section
                if not df_h_action.empty:
                    rows, total, active_days = generate_report.build_section_rows(df_h_action, icols, day_cols, date_col)
                    sections.append(('ដល់ប៉ុស្តិ៍ (Action Needed)', rows, total, icols, active_days))
                    counts['Pending'] += total
                    overall['Pending'] += total
            else:
                icols = generate_report.REPORT_COLS[rn]
                rows, total, active_days = generate_report.build_section_rows(df_h, icols, day_cols, date_col)
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
            rn_file = rn
            if '/' in rn_file or '(' in rn_file or ' ' in rn_file:
                rn_file = str(rn_file).replace(' / ', '_').replace(' ', '_').replace('(', '').replace(')', '')
            tmp_xlsx = os.path.join(output_dir, f"Report_{handle}_{rn_file}_{datetime.now().strftime('%d_%m_%Y_%H%M%S')}.xlsx")
            generate_report.build_handle_excel(
                f"Report_{handle}_{rn}", [(rn, rows, total, icols, active_days)], 
                day_cols, excel_design, tmp_xlsx, mode=mode, order_created_map=order_created_map
            )
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

    # Save final sheet
    final_xlsx = os.path.join(output_dir, f"Report_All_{datetime.now().strftime('%d_%m_%Y')}.xlsx")
    if all_handle_sections:
        # Override build_final_excel to support the dynamic tabs
        from collections import defaultdict
        wb = Workbook()
        report_types = ['Pickup', 'Delivery', 'កំពុងលើផ្លូវ (Transit)', 'ដល់ប៉ុស្តិ៍ (Action Needed)']
        
        shared_days = sorted(set(
            d for _, sections in all_handle_sections
            for _, _, _, _, ad in sections
            for d in ad
        ))

        combined_data = defaultdict(list)
        index_cols_map = {}
        for handle, sections in all_handle_sections:
            for rn, rows, total, index_cols, active_days in sections:
                index_cols_map[rn] = index_cols
                non_footer_rows = [r for r in rows if r.get(index_cols[0]) != 'Grand Total']
                combined_data[rn].extend(non_footer_rows)

        for idx, rn in enumerate(report_types):
            if idx == 0:
                ws = wb.active
                ws.title = rn
            else:
                ws = wb.create_sheet(title=rn)
                
            rows = combined_data[rn]
            icols = index_cols_map.get(rn)
            if not icols:
                icols = generate_report.REPORT_COLS[rn]

            if not rows:
                footer = {col: '' for col in icols}
                footer[icols[0]] = 'Grand Total'
                footer['Grand Total'] = 0
                rows = [footer]
                combined_active_days = []
            else:
                sort_keys = [c for c in ['ZONE', 'POST OFFICE HANDLE', 'CURRENT POST OFFICE', 'ORDER ID'] if c in icols]
                def get_sort_key(row):
                    return tuple(str(row.get(k, '') or '').strip().upper() for k in sort_keys)
                rows = sorted(rows, key=get_sort_key)
                combined_active_days = shared_days
                col_totals = defaultdict(int)
                grand_total = 0
                for row in rows:
                    grand_total += 1
                    for d in combined_active_days:
                        val = row.get(d)
                        if isinstance(val, (int, float)) and val != '':
                            col_totals[d] += int(val)
                footer = {col: '' for col in icols}
                footer[icols[0]] = 'Grand Total'
                for d in combined_active_days:
                    footer[d] = col_totals[d] if col_totals[d] > 0 else ''
                footer['Grand Total'] = grand_total
                rows.append(footer)

            generate_report._write_table(
                ws, 1, 1, rn, rows, icols, combined_active_days, excel_design,
                handle="ALL BRANCHES", show_top_title=False, max_index=len(icols),
                order_created_map=order_created_map
            )
            generate_report._set_col_widths(ws)
            
        wb.save(final_xlsx)
    else:
        wb = Workbook()
        wb.save(final_xlsx)

    if return_metadata:
        return {
            'overall_counts': overall,
            'handle_results': handle_results,
            'final_xlsx':     final_xlsx,
        }
    return handle_results

from datetime import timedelta
print("Running Khmer styled report generation...")
result = generate_reports_from_data_khmer(
    "test_detail.xlsx", "post_office_lookup.csv", "test_out_khmer", return_metadata=True, mode="wide"
)

# Render one of the split Pending files (e.g. Action Needed)
first_hr = result["handle_results"][0]
print(f"First handle: {first_hr['handle']}")
print("Generated files:")
for f in first_hr["handle_files"]:
    print(f"  Path: {f['path']}")

# Let's render the 'កំពុងលើផ្លូវ (Transit)' or 'ដល់ប៉ុស្តិ៍ (Action Needed)' sheet
transit_file = next((f['path'] for f in first_hr["handle_files"] if "Transit" in f['path']), None)
action_file = next((f['path'] for f in first_hr["handle_files"] if "Action" in f['path']), None)

if action_file:
    print(f"Rendering Action Needed: {action_file}")
    img_buf = excel_to_image.excel_to_image(action_file)
    out_png = "c:/Users/DELL/.gemini/antigravity/brain/b8105454-e4a4-4a86-b723-6ef7dcbedc27/test_action_needed.png"
    with open(out_png, "wb") as f_img:
        f_img.write(img_buf.getvalue())
    print(f"Action needed image successfully rendered to {out_png}")

if transit_file:
    print(f"Rendering Transit: {transit_file}")
    img_buf = excel_to_image.excel_to_image(transit_file)
    out_png = "c:/Users/DELL/.gemini/antigravity/brain/b8105454-e4a4-4a86-b723-6ef7dcbedc27/test_transit.png"
    with open(out_png, "wb") as f_img:
        f_img.write(img_buf.getvalue())
    print(f"Transit image successfully rendered to {out_png}")
