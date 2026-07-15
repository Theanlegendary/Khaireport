import os
import shutil

# Copy original file to temp
original_file = "c:/Users/DELL/Desktop/daily_push/generate_report.py"
temp_file = "c:/Users/DELL/Desktop/daily_push/scratch/generate_report_temp.py"

shutil.copy(original_file, temp_file)
print(f"Copied {original_file} to {temp_file}")

# Read content of temp file
with open(temp_file, "r", encoding="utf-8") as f:
    content = f.read()

# 1. Define new helper functions and maps
khmer_helpers = """
# ── Khmer Helpers ─────────────────────────────────────────────────────────────
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
    'Pending': 'អីវ៉ាន់កំពុងរង់ចាំ (Pending)',
    'Pickup': 'អីវ៉ាន់ត្រូវយក (Pickup)',
    'Delivery': 'អីវ៉ាន់ត្រូវដឹក (Delivery)',
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
"""

# Insert helpers after imports
import_marker = "from openpyxl.utils import get_column_letter"
content = content.replace(import_marker, f"{import_marker}\n{khmer_helpers}")

# 2. Update REPORT_COLS and REPORT_FILTER_COLS definitions
old_report_cols = """REPORT_COLS = {
    'Pickup':   ['ZONE', 'POST OFFICE HANDLE', 'CURRENT POST OFFICE', 'ORDER ID', 'Cus name', 'Phone'],
    'Delivery': ['ZONE', 'POST OFFICE HANDLE', 'CURRENT POST OFFICE', 'ORDER ID', 'RECEIVER'],
    'Pending':  ['ZONE', 'CURRENT POST OFFICE', 'ORDER ID', 'REMARK'],
}"""

new_report_cols = """REPORT_COLS = {
    'Pickup':   ['ZONE', 'POST OFFICE HANDLE', 'CURRENT POST OFFICE', 'ORDER ID', 'Cus name', 'Phone'],
    'Delivery': ['ZONE', 'POST OFFICE HANDLE', 'CURRENT POST OFFICE', 'ORDER ID', 'RECEIVER'],
    'Pending':  ['ZONE', 'CURRENT POST OFFICE', 'ORDER ID', 'REMARK', 'NEXT_ACTION'],
}"""

content = content.replace(old_report_cols, new_report_cols)

# 3. Update _write_table signature to accept order_status_map
old_write_table_sig = """def _write_table(ws, start_row, start_col, report_name, rows, index_cols, active_days, dc, handle="", show_top_title=False, max_index=0, order_created_map=None):"""
new_write_table_sig = """def _write_table(ws, start_row, start_col, report_name, rows, index_cols, active_days, dc, handle="", show_top_title=False, max_index=0, order_created_map=None, order_status_map=None):"""
content = content.replace(old_write_table_sig, new_write_table_sig)

# 4. Translate header names in sheet cell writing
old_write_cell = "ws.cell(r, col_idx, col_name)"
new_write_cell = "ws.cell(r, col_idx, translate_header(col_name))"
content = content.replace(old_write_cell, new_write_cell)

# 5. Row highlight coloring in data row loop
old_overdue_check = """        # Check if this row's order is over 2 days old
        is_overdue = False
        if not is_total and order_created_map:
            order_id = normalize_id(row_dict.get('ORDER ID', ''))
            created_date = order_created_map.get(order_id)
            if created_date and (today - created_date).days > HIGHLIGHT_OVER_DAYS:
                is_overdue = True"""

new_overdue_check = """        # Check if this row's order is over 2 days old
        is_overdue = False
        if not is_total and order_created_map:
            order_id = normalize_id(row_dict.get('ORDER ID', ''))
            created_date = order_created_map.get(order_id)
            if created_date and (today - created_date).days > HIGHLIGHT_OVER_DAYS:
                is_overdue = True

        # Status highlights (super light green for 420/472, super light orange for 430/460/360/500/520/540)
        row_fill = None
        if not is_total and order_status_map:
            order_id = normalize_id(row_dict.get('ORDER ID', ''))
            sc = order_status_map.get(order_id, '')
            if sc in ('420', '472'):
                row_fill = _fill('E2EFDA')  # Super light green
            elif sc in ('430', '460', '360', '500', '520', '540'):
                row_fill = _fill('FCE4D6')  # Super light orange

        if not row_fill and is_overdue:
            row_fill = highlight_fill"""

content = content.replace(old_overdue_check, new_overdue_check)

# 6. Apply cell styling using row_fill
old_fill_overdue = """            elif is_overdue:
                # Highlight entire row with light red for overdue orders
                cell.fill      = highlight_fill
                cell.font      = _font(fn, '000000', bold=False)
                cell.alignment = _align('center')"""

new_fill_overdue = """            elif row_fill:
                cell.fill      = row_fill
                cell.font      = _font(fn, '000000', bold=False)
                cell.alignment = _align('center')"""

content = content.replace(old_fill_overdue, new_fill_overdue)

# 7. Use Khmer font for NEXT_ACTION and REMARK in sheet cell formatting
old_format_cell = """            elif col_name in index_cols:
                # All other index columns (CURRENT POST OFFICE, ORDER ID, etc.) → center
                cell.font      = _font(fn, '1E293B', bold=False)
                cell.alignment = _align('center')"""

new_format_cell = """            elif col_name in index_cols:
                # All other index columns (CURRENT POST OFFICE, ORDER ID, etc.) → center
                f_name = 'Khmer OS Battambang' if col_name in ('REMARK', 'NEXT_ACTION') else fn
                cell.font      = _font(f_name, '1E293B', bold=False, size=9 if col_name in ('REMARK', 'NEXT_ACTION') else 10)
                cell.alignment = _align('center')"""

content = content.replace(old_format_cell, new_format_cell)

# 8. Update build_handle_excel signature & calls (Wide / Long mode)
old_build_handle_sig = """def build_handle_excel(handle, sections, day_cols, dc, out_path, mode='wide', order_created_map=None):"""
new_build_handle_sig = """def build_handle_excel(handle, sections, day_cols, dc, out_path, mode='wide', order_created_map=None, order_status_map=None):"""
content = content.replace(old_build_handle_sig, new_build_handle_sig)

old_call_write_table_wide = """            _, next_col = _write_table(ws, 1, cur_col, report_name,
                                       rows, index_cols, shared_days, dc, handle=handle, max_index=max_index, order_created_map=order_created_map)"""

new_call_write_table_wide = """            _, next_col = _write_table(ws, 1, cur_col, report_name,
                                       rows, index_cols, shared_days, dc, handle=handle, max_index=max_index, order_created_map=order_created_map, order_status_map=order_status_map)"""

content = content.replace(old_call_write_table_wide, new_call_write_table_wide)

old_call_write_table_long = """            next_row, _ = _write_table(ws, r, 1, report_name,
                                       rows, index_cols, shared_days, dc, handle=handle, show_top_title=(i==0), max_index=max_index, order_created_map=order_created_map)"""

new_call_write_table_long = """            next_row, _ = _write_table(ws, r, 1, report_name,
                                       rows, index_cols, shared_days, dc, handle=handle, show_top_title=(i==0), max_index=max_index, order_created_map=order_created_map, order_status_map=order_status_map)"""

content = content.replace(old_call_write_table_long, new_call_write_table_long)

# 9. Update build_final_excel signature & calls
old_build_final_sig = """def build_final_excel(all_handle_sections, day_cols, dc, out_path, mode='wide', order_created_map=None):"""
new_build_final_sig = """def build_final_excel(all_handle_sections, day_cols, dc, out_path, mode='wide', order_created_map=None, order_status_map=None):"""
content = content.replace(old_build_final_sig, new_build_final_sig)

old_final_excel_write = """        _write_table(ws, 1, 1, rn, rows, icols, combined_active_days, dc,
                     handle="ALL BRANCHES", show_top_title=False, max_index=len(icols),
                     order_created_map=order_created_map)"""

new_final_excel_write = """        _write_table(ws, 1, 1, rn, rows, icols, combined_active_days, dc,
                     handle="ALL BRANCHES", show_top_title=False, max_index=len(icols),
                     order_created_map=order_created_map, order_status_map=order_status_map)"""

content = content.replace(old_final_excel_write, new_final_excel_write)

# 10. Update generate_reports_from_data: extract order_status_map and add NEXT_ACTION to Pending class
old_generate_date_mapping = """    order_created_map = {}
    if date_col and 'ORDER ID' in df.columns:
        parsed_created = pd.to_datetime(df[date_col], dayfirst=True, format='mixed', errors='coerce')
        for order_id, dt in zip(df['ORDER ID'], parsed_created):
            if pd.notna(dt):
                order_created_map[normalize_id(order_id)] = dt.date()"""

new_generate_date_mapping = """    order_created_map = {}
    order_status_map = {}
    if 'ORDER ID' in df.columns:
        if date_col:
            parsed_created = pd.to_datetime(df[date_col], dayfirst=True, format='mixed', errors='coerce')
            for order_id, dt in zip(df['ORDER ID'], parsed_created):
                if pd.notna(dt):
                    order_created_map[normalize_id(order_id)] = dt.date()
        if 'CURRENT STATUS' in df.columns:
            for order_id, status_val in zip(df['ORDER ID'], df['CURRENT STATUS']):
                if pd.notna(order_id) and pd.notna(status_val):
                    sc_val = str(status_val).strip()
                    import re
                    match = re.match(r'^(\\d{3})', sc_val)
                    if match:
                        order_status_map[normalize_id(order_id)] = match.group(1)"""

content = content.replace(old_generate_date_mapping, new_generate_date_mapping)

old_pending_mapping = """        # Add REMARK column for Pending and remove completed statuses
        if rn == 'Pending' and 'STATUS_CODE' in df_t.columns:
            def _pending_remark(row):
                sc = str(row.get('STATUS_CODE', '')).strip()
                return PENDING_REMARK_MAP.get(sc, 'Unknown')
            df_t['REMARK'] = df_t.apply(_pending_remark, axis=1)
        type_data[rn] = df_t"""

new_pending_mapping = """        # Add REMARK column for Pending and remove completed statuses
        if rn == 'Pending' and 'STATUS_CODE' in df_t.columns:
            def _pending_remark(row):
                sc = str(row.get('STATUS_CODE', '')).strip()
                return PENDING_REMARK_MAP.get(sc, 'Unknown')
            def _pending_next_action(row):
                sc = str(row.get('STATUS_CODE', '')).strip()
                return get_next_action(sc)
            df_t['REMARK'] = df_t.apply(_pending_remark, axis=1)
            df_t['NEXT_ACTION'] = df_t.apply(_pending_next_action, axis=1)
        type_data[rn] = df_t"""

content = content.replace(old_pending_mapping, new_pending_mapping)

# In the handle loop, we DO NOT split Pending. We keep it as is!
old_call_build_handle_excel = """            build_handle_excel(f"Report_{handle}_{rn}", [(rn, rows, total, icols, active_days)], day_cols, excel_design, tmp_xlsx, mode=mode, order_created_map=order_created_map)"""
new_call_build_handle_excel = """            build_handle_excel(f"Report_{handle}_{rn}", [(rn, rows, total, icols, active_days)], day_cols, excel_design, tmp_xlsx, mode=mode, order_created_map=order_created_map, order_status_map=order_status_map)"""
content = content.replace(old_call_build_handle_excel, new_call_build_handle_excel)

old_call_build_final_excel = """        build_final_excel(all_handle_sections, day_cols, excel_design, final_xlsx, mode=mode, order_created_map=order_created_map)"""
new_call_build_final_excel = """        build_final_excel(all_handle_sections, day_cols, excel_design, final_xlsx, mode=mode, order_created_map=order_created_map, order_status_map=order_status_map)"""
content = content.replace(old_call_build_final_excel, new_call_build_final_excel)

# Save changes back to temp file
with open(temp_file, "w", encoding="utf-8") as f:
    f.write(content)
print(f"Successfully modified temp file: {temp_file}")
