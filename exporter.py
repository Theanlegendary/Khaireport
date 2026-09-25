"""
exporter.py
Exports exactly 5 report images from the calculated Excel Master Daily Report.
Directly copies ranges via Excel COM CopyPicture -> Chart -> PNG.
No synthetic tables or pillow generation.

Template: 00.Master Daily Report - new - 1509.xlsx (45MB, Sept version)
Sheet layout:
  Province_Report: Image1 at A5:W30, Image5 at A63:O87
  SP_RP:           Image2 at B3:U44
  Agent_RP:        Image3 at A3:P28
  Showroom_RP:     Image4 at A3:P28
"""

import os
import sys
import time
import win32com.client
import pythoncom
from PIL import Image


def export_range_to_image(ws, cell_range, out_path, hide_cols=None, widen_cols=None, scale=2.0):
    """
    Exports a specific worksheet range to PNG using Excel COM CopyPicture.
    """
    ws.Activate()
    time.sleep(0.1)

    # Disable gridlines for clean export
    try:
        ws.Parent.Windows(1).DisplayGridlines = False
    except Exception:
        pass

    # Hide requested columns
    if hide_cols:
        for c in hide_cols:
            try:
                ws.Columns(c).Hidden = True
            except Exception as e:
                print(f"[EXPORTER] Error hiding column {c}: {e}")

    # Widen columns if needed to prevent '###' overflow
    orig_widths = {}
    if widen_cols:
        for col_name, min_w in widen_cols.items():
            try:
                col_obj = ws.Columns(col_name)
                orig_widths[col_name] = col_obj.ColumnWidth
                col_obj.ColumnWidth = max(col_obj.ColumnWidth, min_w)
            except Exception as e:
                print(f"[EXPORTER] Error widening column {col_name}: {e}")

    rng = ws.Range(cell_range)
    time.sleep(0.15)

    # CopyPicture (xlScreen=1, xlBitmap=-4147)
    copied = False
    for attempt in range(5):
        try:
            rng.CopyPicture(1, -4147)
            copied = True
            break
        except Exception:
            time.sleep(0.25)

    if not copied:
        raise RuntimeError(f"Failed to CopyPicture on {ws.Name}!{cell_range}")

    w = rng.Width * scale
    h = rng.Height * scale

    co = ws.ChartObjects().Add(Left=0, Top=0, Width=w, Height=h)
    co.Activate()
    ch = co.Chart
    time.sleep(0.1)
    ch.Paste()

    if ch.Shapes.Count > 0:
        s = ch.Shapes(1)
        s.Left = 0
        s.Top = 0
        s.ScaleWidth(scale, 1)
        s.ScaleHeight(scale, 1)
        co.Width = s.Width
        co.Height = s.Height

    ch.ChartArea.Format.Line.Visible = 0
    ch.ChartArea.Format.Fill.Visible = 0

    abs_out_path = os.path.abspath(out_path)
    os.makedirs(os.path.dirname(abs_out_path), exist_ok=True)
    ch.Export(abs_out_path, "PNG")
    co.Delete()

    # Restore unhidden columns
    if hide_cols:
        for c in hide_cols:
            try:
                ws.Columns(c).Hidden = False
            except Exception:
                pass

    # Restore column widths
    if orig_widths:
        for col_name, orig_w in orig_widths.items():
            try:
                ws.Columns(col_name).ColumnWidth = orig_w
            except Exception:
                pass

    if os.path.exists(abs_out_path):
        with Image.open(abs_out_path) as im:
            print(f"[EXPORTER] Generated {os.path.basename(abs_out_path)} ({im.size[0]}x{im.size[1]})")
        return abs_out_path

    raise FileNotFoundError(f"Exported image not found at {abs_out_path}")


def export_5_report_images(xlsx_path, out_dir):
    """
    Opens the workbook and exports all 5 management report images.

    Template: 00.Master Daily Report - new - 1509.xlsx (Sept version, 45MB)

    1. BILL ORDER - DAY:       Province_Report!A5:X30   (22 branches, no hidden cols)
    2. [SERVICE POINT] REPORT: SP_RP!B3:U44             (36 SPs by province)
    3. [AGENT] REPORT:         Agent_RP!A3:P28          (22 branches)
    4. [SHOWROOM] REPORT:      Showroom_RP!A3:P28       (22 branches)
    5. CUSTOMER REPORT:        Province_Report!A63:O87  (New/Order/Record by 22 branches)

    Returns dict with keys 'img1', 'img2', 'img3', 'img4', 'img5'
    """
    abs_xlsx = os.path.abspath(xlsx_path)
    os.makedirs(out_dir, exist_ok=True)

    pythoncom.CoInitialize()
    excel = None
    wb = None

    results = {}

    try:
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
        except Exception:
            excel = win32com.client.Dispatch("Excel.Application")

        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False

        print(f"[EXPORTER] Opening Excel for export: {abs_xlsx}")
        try:
            wb = excel.Workbooks.Open(abs_xlsx)
        except Exception:
            wb = excel.Workbooks.Open(abs_xlsx, 0, False, 5, '', '', True, 1, '', True, False, 0, False, 1, 1)

        # ----------------------------------------------------------------
        # 1. BILL ORDER - DAY: Province_Report!A5:X30
        #    Rows: 5-7=headers, 8=MEC, 9-30=22 branches
        #    Cols: A-X = 22 branches metrics
        # ----------------------------------------------------------------
        ws_p = wb.Worksheets("Province_Report")
        p1 = os.path.join(out_dir, "1_day_report.png")
        export_range_to_image(ws_p, "A5:X30", p1)
        results["img1"] = p1

        # ----------------------------------------------------------------
        # 2. SERVICE POINT REPORT: SP_RP!B3:U44
        #    Rows: 3=title, 4-5=headers, 6=MEC, 7-44=36 SPs by province
        #    Cols: B-U
        # ----------------------------------------------------------------
        ws_sp = wb.Worksheets("SP_RP")
        p2 = os.path.join(out_dir, "2_sp_order_express.png")
        export_range_to_image(ws_sp, "B3:U44", p2)
        results["img2"] = p2

        # ----------------------------------------------------------------
        # 3. AGENT REPORT: Agent_RP!A3:P28
        #    Rows: 3=title, 4-5=headers, 6=MEC, 7-28=22 branches
        #    Cols: A=NO, B=Branch Code, C=Branch, D=No.Agents, E-I=In month, J-N=Inday, O-P=Delta
        # ----------------------------------------------------------------
        ws_ar = wb.Worksheets("Agent_RP")
        p3 = os.path.join(out_dir, "3_agent_report.png")
        export_range_to_image(ws_ar, "A3:P28", p3)
        results["img3"] = p3

        # ----------------------------------------------------------------
        # 4. SHOWROOM REPORT: Showroom_RP!A3:P28
        #    Same layout as Agent_RP
        # ----------------------------------------------------------------
        ws_sr = wb.Worksheets("Showroom_RP")
        p4 = os.path.join(out_dir, "4_showroom_report.png")
        export_range_to_image(ws_sr, "A3:P28", p4)
        results["img4"] = p4

        # ----------------------------------------------------------------
        # 5. CUSTOMER REPORT: Province_Report!A63:O87
        #    Rows: 63-64=headers, 65=MEC, 66-87=22 branches
        #    Cols: A-C=identity, D-G=New customer, H-L=Order of new Customer, M-O=Record new Customer
        # ----------------------------------------------------------------
        p5 = os.path.join(out_dir, "5_customer_report.png")
        export_range_to_image(ws_p, "A63:O87", p5)
        results["img5"] = p5

        return results

    finally:
        if wb:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        if excel:
            try:
                excel.Quit()
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
