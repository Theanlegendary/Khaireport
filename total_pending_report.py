"""
total_pending_report.py — TOTAL PENDING Report Generator
==========================================================
Generates CEO executive summary and audit dataset for all pending shipments
across all branches, grouped by Branch, Facility Type (P, S, A), and Age (Days).
Designed with the official CEO Executive Teal Table aesthetic and intuitive color tiers.

SIMPLIFIED - NO GRACE PERIOD LOGIC (like /penalty)

Approved Statuses (14):  # REMOVED 420 and 472
- NOT ASSIGN / Branch (3): 306, 309, 400
- DELIVERY (11): 401, 402, 430, 460, 470, 471, 480, 500, 510, 511, 512

Age Buckets (ACTUAL HOURS - NO GRACE ADJUSTMENTS):
- 0 Days:   0:00 -> 23:59 (< 24.0 hours)
- 1 Day:   24:00 -> 47:59 (>= 24.0 and < 48.0 hours)
- 2 Days:  48:00 -> 71:59 (>= 48.0 and < 72.0 hours)
- 3 Days:  72:00 -> 95:59 (>= 72.0 and < 96.0 hours)
- 4 Days:  96:00 -> 119:59 (>= 96.0 and < 120.0 hours)
- 5 Days: 120:00 -> 167:59 (>= 120.0 and < 168.0 hours, covers 5 to 7 days)
- > 7 Days: >= 168.0 and < 720.0 hours (> 7 Days and < 30 Days, limited to this month)
- Over 30 Days: Excluded (>= 720.0 hours)

NO MORE:
- No status 420/472 grace period
- No checking tracking history logs
- Uses actual aging hours directly (cleaner and faster like /penalty)
"""

import os
import io
import re
import json
import tempfile
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor
import requests
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import excel_to_image

APPROVED_STATUSES = {
    '306', '309', '400',
    '401', '402', '430', '460', '470', '471', '480', '500', '510', '511', '512'
    # REMOVED: '420', '472' - no longer check these statuses or their history
}

FACILITY_COLS = ['Servicepoint', 'Showroom', 'Agent']
DAY_COLS = ['0 Days', '1 Day', '2 Days', '3 Days', '4 Days', '5 Days', '> 7 Days']
DELAY_COLS = ['3 Days', '4 Days', '5 Days', '> 7 Days']
DELAY_COL_NAME = 'Total >= 3 Days'

TS_COLS_PRIORITY = [
    'CURRENT TIME',
    'STATUS 306 AT STORE / AGENT (LAST TIME)',
    'STATUS 306 AT STORE / AGENT FROM HUB (FIRST TIME)',
    'STATUS 302/310 AT RECEIVING STORE / RECEIVING AGENT (FIRST TIME)',
    'STATUS 306  AT ORIGIN HUB (FIRST TIME)',
    'STATUS 210 TIME',
    'CREATED DATE'
]


def get_branch_code(po_code: str) -> str:
    """
    Map post office / showroom / agent code to parent Branch code.
    Only 'PNP' retains 'P'; all other provincial branches have 3 letters with no 'P'
    (e.g., BAN, BAT, CHA, KAN, KAM, KOH, KRA, MON, ODD, PRE, PUR, SIE, SIH, SPE, TAK, etc.).
    """
    c = str(po_code or '').strip().upper()
    if not c or c == 'NAN':
        return 'UNKNOWN'
    if c.startswith('PNP'):
        return 'PNP'
    return c[:3]


def get_facility_type(po_code: str) -> str:
    """Classify facility as Servicepoint, Showroom, or Agent."""
    c = str(po_code or '').strip().upper()
    if len(c) >= 4:
        ch = c[3]
        if ch == 'S':
            return 'Showroom'
        elif ch == 'A':
            return 'Agent'
        elif ch == 'P':
            return 'Servicepoint'
    return 'Servicepoint'

# REMOVED - No longer needed (no grace period logic)
# def fetch_pending_history_allowances(order_ids: list[str], bearer_token: str = None) -> dict[str, dict[str, bool]]:

# REMOVED - No longer needed (no grace period logic)
# def fetch_pending_history_allowances(order_ids: list[str], bearer_token: str = None) -> dict[str, dict[str, bool]]:
#     """
#     Check real-time tracking trips for candidate pending bills (>= 24h old).
#     Returns a mapping of order_id -> {'has_420': bool, 'has_472': bool}
#     indicating whether status 420 or 472 appeared anywhere in the tracking history log.
#     """
#     [FUNCTION REMOVED - No longer checking history for 420/472]


def process_pending_data(src_path_or_df):
    """
    Load and process TMS data to build the TOTAL PENDING report dataset.
    Returns:
        (summary_df, grand_total_dict, df_pending_detail)
    """
    if isinstance(src_path_or_df, pd.DataFrame):
        df = src_path_or_df.copy()
    else:
        try:
            df = pd.read_excel(src_path_or_df, engine='calamine')
        except Exception:
            df = pd.read_excel(src_path_or_df)

    # Standardize column names
    df.columns = [str(c).strip() for c in df.columns]

    # Extract 3-digit status code
    sc = df['CURRENT STATUS'].astype(str).str.extract(r'^(\d{3})')[0]
    df_p = df[sc.isin(APPROVED_STATUSES)].copy()
    df_p['STATUS_CODE'] = sc[sc.isin(APPROVED_STATUSES)]

    # Exclude test orders & test post offices (e.g. BANCHI_TEST, TEST PH NOM PENH, test orders)
    test_col = next((c for c in df_p.columns if str(c).strip().lower() in ('is test', 'đơn test', 'don test')), None)
    if test_col:
        df_p = df_p[df_p[test_col].isna() | df_p[test_col].astype(str).str.strip().isin(['', 'nan', 'NaN', '#N/A'])].copy()

    # Exclude any row where 'test' appears in key text fields
    check_cols = [c for c in df_p.columns if any(k in str(c).upper() for k in ('ORDER', 'STATUS', 'POST', 'OFFICE', 'SENDER', 'RECEIVER', 'NOTE', 'REMARK', 'GOODS', 'DESC', 'PRODUCT'))]
    if check_cols:
        test_mask = df_p[check_cols].astype(str).apply(lambda s: s.str.lower().str.contains('test', na=False)).any(axis=1)
        df_p = df_p[~test_mask].copy()

    # Exclude test bills registered in test_bills.txt, delayed_bills.json, test.xlsx, etc.
    try:
        from penalty_report import load_test_bills
        test_bill_ids = load_test_bills()
    except Exception:
        test_bill_ids = set()

    if test_bill_ids:
        order_col = next(
            (c for c in df_p.columns if str(c).strip().upper() in ('ORDER ID', 'ORDER_ID', 'BILL_ID', 'BILL ID', 'WAYBILL', 'ORDER ID/ WAYBILL')),
            'ORDER ID' if 'ORDER ID' in df_p.columns else None
        )
        if order_col:
            clean_ids = df_p[order_col].astype(str).str.strip().str.upper().str.replace(r'\.0$', '', regex=True)
            df_p = df_p[~clean_ids.isin(test_bill_ids)].copy()

    # Deduplicate by ORDER ID (keep latest record)
    if 'ORDER ID' in df_p.columns:
        df_p = df_p.drop_duplicates(subset=['ORDER ID'], keep='last').copy()

    # Exclude central MEGA/HUB/DVC sorting facilities (not branch delivery points)
    is_hub = df_p['CURRENT POST OFFICE'].astype(str).str.contains('MEGA|HUB|DVC', case=False, na=False)
    df_branch = df_p[~is_hub].copy()

    # ── PURGE DELIVERED / SHIPPED BILLS (LIVE TRACKING CROSS-CHECK) ───────────
    try:
        from shipped_filter import filter_shipped_bills_from_df
        df_branch, removed_shipped = filter_shipped_bills_from_df(df_branch, verify_live=True)
    except Exception as e_shipped:
        print(f"[TOTAL_PENDING] Warning: Live shipped verification error: {e_shipped}")

    df_branch['Branch'] = df_branch['CURRENT POST OFFICE'].apply(get_branch_code)
    df_branch['Facility_Type'] = df_branch['CURRENT POST OFFICE'].apply(get_facility_type)

    # Calculate timestamps and actual age
    now = datetime.now()

    def _calc_actual_ts(row):
        ts = None
        for c in TS_COLS_PRIORITY:
            if c in row and pd.notna(row[c]):
                val = str(row[c]).strip()
                if val and val.lower() != 'nan':
                    dt = pd.to_datetime(val, dayfirst=True, format='mixed', errors='coerce')
                    if pd.notna(dt):
                        ts = dt
                        break
        if ts is None:
            actual_hours = 0.0
            ts_str = ''
        else:
            actual_hours = max(0.0, (now - ts).total_seconds() / 3600.0)
            ts_str = ts.strftime('%d/%m/%Y %H:%M')
        return pd.Series([ts_str, round(actual_hours, 1)])

    ts_df = df_branch.apply(_calc_actual_ts, axis=1)
    df_branch['History_Timestamp'] = ts_df[0]
    df_branch['Actual_Hours'] = ts_df[1]

    # NO MORE GRACE PERIOD OR HISTORY CHECKING - SIMPLIFIED LIKE /PENALTY
    def _calc_aging_bucket(row):
        actual_hours = row['Actual_Hours']
        
        # NO GRACE ADJUSTMENTS - USE ACTUAL HOURS DIRECTLY
        adjusted_hours = actual_hours

        # Bucket classification based on ACTUAL hours (no grace deductions)
        if adjusted_hours >= 720.0:
            bucket = 'EXCLUDED_OVER_30'
        elif adjusted_hours < 24.0:
            bucket = '0 Days'
        elif adjusted_hours < 48.0:
            bucket = '1 Day'
        elif adjusted_hours < 72.0:
            bucket = '2 Days'
        elif adjusted_hours < 96.0:
            bucket = '3 Days'
        elif adjusted_hours < 120.0:
            bucket = '4 Days'
        elif adjusted_hours < 168.0:
            bucket = '5 Days'
        else:
            bucket = '> 7 Days'

        return pd.Series([round(adjusted_hours, 1), bucket])  # Only return hours and bucket

    aging_df = df_branch.apply(_calc_aging_bucket, axis=1)
    df_branch['Adjusted_Hours'] = aging_df[0]
    df_branch['Age_Bucket'] = aging_df[1]
    # REMOVED: Grace_Note column completely - no longer needed

    # Exclude orders >= 30 days (older than 30 days / not in this month)
    df_branch = df_branch[df_branch['Age_Bucket'] != 'EXCLUDED_OVER_30'].copy()

    # Aggregate by branch
    unique_branches = sorted(df_branch['Branch'].unique(), key=lambda x: (0 if x == 'PNP' else 1, x))
    rows = []

    for b in unique_branches:
        sub = df_branch[df_branch['Branch'] == b]
        sp_cnt = int((sub['Facility_Type'] == 'Servicepoint').sum())
        sr_cnt = int((sub['Facility_Type'] == 'Showroom').sum())
        ag_cnt = int((sub['Facility_Type'] == 'Agent').sum())
        tot = len(sub)
        day_counts = {d: int((sub['Age_Bucket'] == d).sum()) for d in DAY_COLS}

        # Reconciliation check
        assert sp_cnt + sr_cnt + ag_cnt == tot, f"Facility mismatch for {b}"
        assert sum(day_counts.values()) == tot, f"Age sum mismatch for {b}"

        delay_cnt = sum(day_counts[d] for d in DELAY_COLS)

        row_data = {
            'Branch': b,
            'Servicepoint': sp_cnt,
            'Showroom': sr_cnt,
            'Agent': ag_cnt
        }
        row_data.update(day_counts)
        row_data[DELAY_COL_NAME] = delay_cnt
        row_data['Total'] = tot
        rows.append(row_data)

    summary_df = pd.DataFrame(rows)

    # Grand total calculation
    grand_total = {
        'Branch': 'TOTAL',
        'Servicepoint': int(summary_df['Servicepoint'].sum()) if not summary_df.empty else 0,
        'Showroom': int(summary_df['Showroom'].sum()) if not summary_df.empty else 0,
        'Agent': int(summary_df['Agent'].sum()) if not summary_df.empty else 0,
    }
    for d in DAY_COLS:
        grand_total[d] = int(summary_df[d].sum()) if not summary_df.empty else 0
    grand_total[DELAY_COL_NAME] = int(summary_df[DELAY_COL_NAME].sum()) if not summary_df.empty else 0
    grand_total['Total'] = int(summary_df['Total'].sum()) if not summary_df.empty else 0

    # Backwards compatibility aliases
    grand_total['P'] = grand_total['Servicepoint']
    grand_total['S'] = grand_total['Showroom']
    grand_total['A'] = grand_total['Agent']

    return summary_df, grand_total, df_branch


def export_total_pending_excel(summary_df, grand_total, df_detail, out_xlsx_path):
    """
    Exports clean two-sheet Excel file matching the CEO Executive Teal Table style
    with color-coded numbers:
    1. 'Total Pending Summary' — executive pivot table with teal header, color-coded numbers, and light-teal total
    2. 'Pending Details' — audit drill-down records
    """
    wb = openpyxl.Workbook()
    ws_sum = wb.active
    ws_sum.title = "Total Pending Summary"
    ws_sum.views.sheetView[0].showGridLines = True

    # Official CEO Colors & Fonts
    header_fill = PatternFill(start_color="2E8B8B", end_color="2E8B8B", fill_type="solid")  # Teal
    total_fill  = PatternFill(start_color="B8E6E6", end_color="B8E6E6", fill_type="solid")  # Light Teal
    white_fill  = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")  # Pure White

    title_font  = Font(name="Arial", size=12, bold=True, color="FFFFFF")
    header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    branch_font = Font(name="Arial", size=10, bold=True, color="000000")

    # Dynamic Number Colors for Data Rows
    fn_zero      = Font(name="Arial", size=10, color="94A3B8")              # Dim Gray for zeros
    fn_facility  = Font(name="Arial", size=10, color="002060")              # Dark Blue (Servicepoint, Showroom, Agent)
    fn_normal    = Font(name="Arial", size=10, color="0F172A")              # Normal Dark (0 Days, 1 Day, 2 Days)
    fn_alert     = Font(name="Arial", size=10, bold=True, color="FF0000")   # Red (3 Days, 4 Days, 5 Days, > 7 Days)
    fn_col_total = Font(name="Arial", size=10.5, bold=True, color="002060") # Bold Dark Blue (Total column)

    # Grand Total Row Fonts
    fn_tot_label    = Font(name="Arial", size=11, bold=True, color="000000")
    fn_tot_facility = Font(name="Arial", size=11, bold=True, color="002060") # Dark Blue
    fn_tot_normal   = Font(name="Arial", size=11, bold=True, color="0F172A") # Dark
    fn_tot_alert    = Font(name="Arial", size=11, bold=True, color="FF0000") # Red
    fn_tot_total    = Font(name="Arial", size=11.5, bold=True, color="002060") # Bold Dark Blue

    border_style = Border(
        left=Side(style="thin", color="CCCCCC"),
        right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),
        bottom=Side(style="thin", color="CCCCCC")
    )

    columns = ['Branch'] + FACILITY_COLS + DAY_COLS + [DELAY_COL_NAME, 'Total']

    # 1. Title Banner
    ws_sum.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(columns))
    now_str = datetime.now().strftime("%d/%m/%Y %H:%M")
    title_cell = ws_sum.cell(row=1, column=1, value=f"TOTAL PENDING REPORT (ALL BRANCHES)  —  {now_str}")
    title_cell.font = title_font
    title_cell.fill = header_fill
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    title_cell.border = border_style
    ws_sum.row_dimensions[1].height = 28

    # 2. Table Headers
    ws_sum.row_dimensions[2].height = 28
    for c_idx, col_name in enumerate(columns, 1):
        cell = ws_sum.cell(row=2, column=c_idx, value=col_name)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border_style

    # 3. Data Rows with Colored Numbers
    start_data_row = 3
    for r_idx, row in summary_df.iterrows():
        cur_r = start_data_row + r_idx
        ws_sum.row_dimensions[cur_r].height = 22
        for c_idx, col_name in enumerate(columns, 1):
            val = row[col_name]
            cell = ws_sum.cell(row=cur_r, column=c_idx, value=val)
            cell.fill = white_fill
            cell.border = border_style
            if col_name == 'Branch':
                cell.font = branch_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = '#,##0'
                if val == 0:
                    cell.font = fn_zero
                elif col_name in ('Servicepoint', 'Showroom', 'Agent'):
                    cell.font = fn_facility
                elif col_name in ('0 Days', '1 Day', '2 Days'):
                    cell.font = fn_normal
                elif col_name in ('3 Days', '4 Days', '5 Days', '> 7 Days', DELAY_COL_NAME):
                    cell.font = fn_alert
                elif col_name == 'Total':
                    cell.font = fn_col_total

    # 4. Total Row with Excel SUM formulas and Colored Numbers
    end_data_row = start_data_row + len(summary_df) - 1
    tot_r = end_data_row + 1
    ws_sum.row_dimensions[tot_r].height = 26

    tot_label = ws_sum.cell(row=tot_r, column=1, value="TOTAL")
    tot_label.font = fn_tot_label
    tot_label.fill = total_fill
    tot_label.border = border_style
    tot_label.alignment = Alignment(horizontal="left", vertical="center")

    for c_idx in range(2, len(columns) + 1):
        col_letter = get_column_letter(c_idx)
        col_name = columns[c_idx - 1]
        tot_val = grand_total.get(col_name, 0)
        cell = ws_sum.cell(row=tot_r, column=c_idx, value=tot_val)
        cell.fill = total_fill
        cell.border = border_style
        cell.alignment = Alignment(horizontal="right", vertical="center")
        cell.number_format = '#,##0'

        if col_name in ('Servicepoint', 'Showroom', 'Agent'):
            cell.font = fn_tot_facility
        elif col_name in ('0 Days', '1 Day', '2 Days'):
            cell.font = fn_tot_normal
        elif col_name in ('3 Days', '4 Days', '5 Days', '> 7 Days', DELAY_COL_NAME):
            cell.font = fn_tot_alert
        elif col_name == 'Total':
            cell.font = fn_tot_total

    # Set Column Widths for Summary
    for c_idx, col_name in enumerate(columns, 1):
        col_let = get_column_letter(c_idx)
        if col_name in ('Branch', 'Total'):
            ws_sum.column_dimensions[col_let].width = 13.0
        elif col_name in ('Servicepoint', 'Showroom', 'Agent'):
            ws_sum.column_dimensions[col_let].width = 14.0
        elif col_name == DELAY_COL_NAME:
            ws_sum.column_dimensions[col_let].width = 18.0
        elif col_name == '> 7 Days':
            ws_sum.column_dimensions[col_let].width = 13.0
        else:
            ws_sum.column_dimensions[col_let].width = 10.5

    # =========================================================================
    # Sheet 2: Pending Details
    # =========================================================================
    ws_det = wb.create_sheet(title="Pending Details")
    ws_det.views.sheetView[0].showGridLines = True

    detail_cols = [
        ('NO', 8),
        ('Branch', 12),
        ('Facility Type', 14),
        ('Current Post Office', 20),
        ('Order ID', 18),
        ('Current Status', 25),
        ('Status Code', 12),
        ('History Timestamp', 20),
        ('Actual Hours', 14),
        ('Adjusted Hours', 15),
        ('Age Bucket', 14),
        # REMOVED: ('Grace Note', 22) - No longer needed (no grace logic)
        ('Sender', 25),
        ('Receiver', 25),
        ('Phone', 16),
        ('Delivery Post Office', 20),
    ]

    ws_det.row_dimensions[1].height = 25
    for c_idx, (col_name, width) in enumerate(detail_cols, 1):
        cell = ws_det.cell(row=1, column=c_idx, value=col_name)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border_style
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws_det.column_dimensions[get_column_letter(c_idx)].width = width

    # Populate details
    det_row = 2
    for _, item in df_detail.iterrows():
        ws_det.row_dimensions[det_row].height = 20
        ws_det.cell(row=det_row, column=1, value=det_row - 1).alignment = Alignment(horizontal='center')
        ws_det.cell(row=det_row, column=2, value=str(item.get('Branch', ''))).alignment = Alignment(horizontal='center')
        ws_det.cell(row=det_row, column=3, value=str(item.get('Facility_Type', ''))).alignment = Alignment(horizontal='center')
        ws_det.cell(row=det_row, column=4, value=str(item.get('CURRENT POST OFFICE', ''))).alignment = Alignment(horizontal='center')
        ws_det.cell(row=det_row, column=5, value=str(item.get('ORDER ID', ''))).alignment = Alignment(horizontal='center')
        ws_det.cell(row=det_row, column=6, value=str(item.get('CURRENT STATUS', ''))).alignment = Alignment(horizontal='left')
        ws_det.cell(row=det_row, column=7, value=str(item.get('STATUS_CODE', ''))).alignment = Alignment(horizontal='center')
        ws_det.cell(row=det_row, column=8, value=str(item.get('History_Timestamp', ''))).alignment = Alignment(horizontal='center')
        ws_det.cell(row=det_row, column=9, value=item.get('Actual_Hours', 0.0)).alignment = Alignment(horizontal='right')
        ws_det.cell(row=det_row, column=10, value=item.get('Adjusted_Hours', 0.0)).alignment = Alignment(horizontal='right')
        ws_det.cell(row=det_row, column=11, value=str(item.get('Age_Bucket', ''))).alignment = Alignment(horizontal='center')
        # REMOVED: Grace_Note column (column 12) - No longer needed
        ws_det.cell(row=det_row, column=12, value=str(item.get('SENDER', ''))).alignment = Alignment(horizontal='left')
        ws_det.cell(row=det_row, column=13, value=str(item.get('RECEIVER', ''))).alignment = Alignment(horizontal='left')
        ws_det.cell(row=det_row, column=14, value=str(item.get('Phone', item.get('PHONE', '')))).alignment = Alignment(horizontal='center')
        ws_det.cell(row=det_row, column=15, value=str(item.get('DELIVERY POST OFFICE', ''))).alignment = Alignment(horizontal='center')

        data_font = Font(name="Arial", size=10, color="000000")
        for col_i in range(1, len(detail_cols) + 1):
            ws_det.cell(row=det_row, column=col_i).font = data_font
            ws_det.cell(row=det_row, column=col_i).border = border_style
        det_row += 1

    wb.save(out_xlsx_path)
    return out_xlsx_path


def render_total_pending_image(summary_df, grand_total, out_png_path=None, df_detail=None, xlsx_path=None):
    """
    Renders the exact CEO Executive Table design directly to image using excel_to_image.
    Produces a crisp, high-definition spreadsheet representation.
    """
    temp_created = False
    if not xlsx_path or not os.path.exists(xlsx_path):
        temp_dir = tempfile.mkdtemp(prefix="pending_render_")
        xlsx_path = os.path.join(temp_dir, "temp_pending.xlsx")
        export_total_pending_excel(summary_df, grand_total, df_detail if df_detail is not None else pd.DataFrame(), xlsx_path)
        temp_created = True

    try:
        img_buf = excel_to_image.excel_to_image(xlsx_path)
        if out_png_path:
            with open(out_png_path, 'wb') as f:
                f.write(img_buf.getvalue())
            img_buf.seek(0)
        return img_buf
    finally:
        if temp_created and os.path.exists(xlsx_path):
            try:
                os.remove(xlsx_path)
                os.rmdir(os.path.dirname(xlsx_path))
            except Exception:
                pass


def format_pending_text_summary(summary_df, grand_total, target_date=None):
    """Formats markdown caption for scheduled or interactive Total Pending reports matching official detail view."""
    if target_date is None:
        target_date = datetime.now()
    stamp_str = target_date.strftime("%d/%m/%Y %H:%M")

    total_val = grand_total.get("Total", 0) if isinstance(grand_total, dict) else (grand_total or 0)
    sp_val = grand_total.get("Servicepoint", 0) if isinstance(grand_total, dict) else 0
    sr_val = grand_total.get("Showroom", 0) if isinstance(grand_total, dict) else 0
    ag_val = grand_total.get("Agent", 0) if isinstance(grand_total, dict) else 0
    delay_val = grand_total.get(DELAY_COL_NAME, grand_total.get("Total >= 3 Days", 0)) if isinstance(grand_total, dict) else 0

    if delay_val == 0 and summary_df is not None and not summary_df.empty:
        if DELAY_COL_NAME in summary_df.columns:
            delay_val = int(summary_df[DELAY_COL_NAME].sum())

    lines = [
        f"TOTAL PENDING REPORT — {stamp_str}",
        f"",
        f"Total Pending: {total_val:,}",
        f"Servicepoint: {sp_val:,} | Showroom: {sr_val:,} | Agent: {ag_val:,}",
        f"Total Delay (>= 3 Days): {delay_val:,}",
    ]

    if summary_df is not None and not summary_df.empty and 'Branch' in summary_df.columns and 'Total' in summary_df.columns:
        lines.append("")
        lines.append("Top 10 Pending Branches:")
        top10 = summary_df.sort_values(by='Total', ascending=False).head(10)
        for idx, (_, row) in enumerate(top10.iterrows(), 1):
            b_code = str(row['Branch'])
            b_tot = int(row['Total'])
            b_delay = int(row.get(DELAY_COL_NAME, row.get('Total >= 3 Days', 0)))
            lines.append(f"{idx}. {b_code}: {b_tot:,} (>=3D: {b_delay:,})")

    return "\n".join(lines)

