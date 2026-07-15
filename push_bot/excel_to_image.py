"""
excel_to_image.py
Renders an openpyxl-formatted Excel sheet to a PNG image using Pillow.
- Day columns: fixed uniform width
- Hidden/empty columns: skipped entirely
- Merged cells: top-left value shown, others blank
"""

import io
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from PIL import Image, ImageDraw, ImageFont

# ── Constants ──────────────────────────────────────────────────────────────────
SCALE       = 2

FONT_SIZE   = 10 * SCALE
ROW_H       = 22 * SCALE    # px — uniform row height
PAD_X       = 8 * SCALE     # horizontal text padding

# Fixed pixel widths per column type
PX_DAY      = 34 * SCALE    # day columns "01"-"31"  — all same size
PX_ZONE     = 72 * SCALE    # ZONE
PX_GT       = 140 * SCALE   # Grand Total
PX_MIN      = 72 * SCALE    # minimum for auto-fit text columns (ORDER ID, names etc.)
PX_MAX      = 230 * SCALE   # maximum for auto-fit
PX_GAP      = 6 * SCALE     # gap/empty separator columns

BG_WHITE    = (255, 255, 255)
BORDER_COL  = (190, 190, 190)
TEXT_DEF    = (0,   0,   0)


def _hex(h):
    h = h.lstrip('#')
    if len(h) == 6:
        return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))
    return TEXT_DEF


def _cell_bg(cell):
    try:
        f = cell.fill
        if f and f.fill_type == 'solid':
            fc = f.fgColor
            if fc.type == 'rgb' and fc.rgb not in ('00000000', 'FFFFFFFF', 'FF000000'):
                return _hex(fc.rgb[-6:])
    except Exception:
        pass
    return None


def _cell_fg(cell):
    try:
        ft = cell.font
        if ft and ft.color and ft.color.type == 'rgb':
            if ft.color.rgb not in ('00000000', 'FF000000'):
                return _hex(ft.color.rgb[-6:])
    except Exception:
        pass
    return TEXT_DEF


def _cell_bold(cell):
    try:
        return bool(cell.font and cell.font.bold)
    except Exception:
        return False


def _cell_align(cell):
    try:
        a = cell.alignment
        if a and a.horizontal in ('center', 'right'):
            return a.horizontal
    except Exception:
        pass
    return 'left'


_WIN_FONTS = "C:/Windows/Fonts"

def _has_khmer(text: str) -> bool:
    """Return True if text contains any Khmer Unicode characters (U+1780–U+17FF)."""
    return any('\u1780' <= ch <= '\u17FF' for ch in text)

def _load_font(size, bold=False):
    # Try system paths first
    for name in (['arialbd.ttf', 'Arial Bold.ttf', 'DejaVuSans-Bold.ttf'] if bold
                 else ['arial.ttf', 'Arial.ttf', 'DejaVuSans.ttf']):
        for prefix in (f"{_WIN_FONTS}/", ""):
            try:
                return ImageFont.truetype(prefix + name, size)
            except Exception:
                pass
    return ImageFont.load_default()

def _load_khmer_font(size):
    """Load Khmer OS Battambang font for Khmer text."""
    for name in ['KhmerOSbattambang.ttf', 'KhmerOSsiemreap.ttf', 'KhmerOScontent.ttf', 'KhmerOS.ttf']:
        for prefix in (f"{_WIN_FONTS}/", ""):
            try:
                return ImageFont.truetype(prefix + name, size)
            except Exception:
                pass
    return _load_font(size, bold=False)

def _get_font(text: str, size: int, bold: bool = False):
    """Return Khmer font if text has Khmer chars, otherwise return normal font."""
    if _has_khmer(text):
        return _load_khmer_font(size + 3) # Khmer fonts are slightly smaller, boost size
    return _load_font(size, bold)


def excel_to_image(xlsx_path: str) -> io.BytesIO:
    # ── Try Excel COM rendering first (for perfect Khmer text shaping and native styling on Windows) ──
    try:
        import win32com.client
        import time
        import os
        from PIL import ImageGrab
        
        abs_path = os.path.abspath(xlsx_path)
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        wb = None
        try:
            wb = excel.Workbooks.Open(abs_path)
            ws = wb.ActiveSheet
            ws.UsedRange.CopyPicture(1, 2)
            time.sleep(0.5) # wait for clipboard
            
            img = None
            for _ in range(5):
                img = ImageGrab.grabclipboard()
                if img:
                    break
                time.sleep(0.5)
                
            if img:
                # Aspect ratio safety check for Telegram
                max_ratio = 18.0
                w, h = img.size
                new_w, new_h = w, h
                if h > 0 and w / h > max_ratio:
                    new_h = int(w / max_ratio)
                elif w > 0 and h / w > max_ratio:
                    new_w = int(h / max_ratio)

                if (new_w, new_h) != (w, h):
                    padded_img = Image.new('RGB', (new_w, new_h), (255, 255, 255))
                    padded_img.paste(img, (0, 0))
                    img = padded_img

                buf = io.BytesIO()
                img.save(buf, format='PNG')
                buf.seek(0)
                return buf
        finally:
            if wb:
                wb.Close(SaveChanges=False)
            excel.Quit()
    except Exception as e:
        # Fall back to Pillow rendering if Excel COM fails or is not available
        pass

    wb = load_workbook(xlsx_path, data_only=True)
    ws = wb.active
    max_row = ws.max_row
    max_col = ws.max_column
    if not max_row or not max_col:
        raise ValueError("Empty sheet")

    # ── 1. Build value grid (handle merged cells) ──────────────────────────────
    grid = [[''] * (max_col + 1) for _ in range(max_row + 1)]
    skip = set()
    for mc in ws.merged_cells.ranges:
        v = ws.cell(mc.min_row, mc.min_col).value
        grid[mc.min_row][mc.min_col] = str(v) if v is not None else ''
        for r in range(mc.min_row, mc.max_row + 1):
            for c in range(mc.min_col, mc.max_col + 1):
                if not (r == mc.min_row and c == mc.min_col):
                    skip.add((r, c))

    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            if (r, c) in skip:
                continue
            try:
                v = ws.cell(r, c).value
            except Exception:
                v = None
            if grid[r][c] == '':
                grid[r][c] = str(v) if v is not None else ''

    # ── 2. Classify columns ────────────────────────────────────────────────────
    # Only scan the first 5 rows to find headers — avoids data rows like
    # "Grand Total" row (which has a number before it) being misclassified
    day_ci  = set()
    zone_ci = set()
    gt_ci   = set()

    HEADER_SCAN_ROWS = min(max_row, 8)
    for r in range(1, HEADER_SCAN_ROWS + 1):
        for c in range(1, max_col + 1):
            val = grid[r][c].strip()
            # Day col: exactly 2-digit 01-31 in a header row
            # Exclude if previous or next cell in same row has content (avoids data Grand Total rows)
            if val.isdigit() and len(val) == 2 and 1 <= int(val) <= 31:
                day_ci.add(c)
            elif val == 'ZONE':
                zone_ci.add(c)
            elif val == 'Grand Total':
                # Only count as GT column if it's a header row (check row has other headers)
                row_vals = [grid[r][cc].strip() for cc in range(1, max_col + 1)]
                if any(v in ('ZONE', 'ORDER ID', 'POST OFFICE HANDLE', 'CURRENT POST OFFICE')
                       for v in row_vals):
                    gt_ci.add(c)

    # Detect empty columns (gap separators between side-by-side tables)
    empty_ci = set(
        c for c in range(1, max_col + 1)
        if not any(grid[r][c].strip() for r in range(1, max_row + 1))
    )

    # ── 3. Compute per-column pixel widths ─────────────────────────────────────
    fn   = _load_font(FONT_SIZE, bold=False)
    fn_b = _load_font(FONT_SIZE, bold=True)
    tmp  = Image.new('RGB', (1, 1))
    d    = ImageDraw.Draw(tmp)

    col_px = []   # pixel width per column (0 = skip/hidden)
    for c in range(1, max_col + 1):
        if c in empty_ci:
            col_px.append(PX_GAP)   # tiny but visible separator
            continue

        # Check if column is hidden in Excel
        letter = get_column_letter(c)
        cd = ws.column_dimensions.get(letter)
        if cd and (cd.hidden or (cd.width is not None and cd.width < 0.5)):
            col_px.append(0)        # fully hidden — skip in image too
            continue

        if c in day_ci:
            col_px.append(PX_DAY)   # ALL day columns same fixed width
        elif c in zone_ci:
            col_px.append(PX_ZONE)
        elif c in gt_ci:
            col_px.append(PX_GT)
        else:
            # Auto-fit text columns
            max_w = PX_MIN
            for r in range(1, max_row + 1):
                text = grid[r][c]
                if not text:
                    continue
                bold = _cell_bold(ws.cell(r, c))
                f = _get_font(text, FONT_SIZE, bold)
                try:
                    bb = d.textbbox((0, 0), text, font=f)
                    w  = bb[2] - bb[0]
                except Exception:
                    w = len(text) * (FONT_SIZE - 2)
                    max_w = max(max_w, w + PAD_X * 2)
            col_px.append(min(max_w, PX_MAX))

    # Columns to actually render (skip 0-width hidden ones)
    render_cols = [c for c in range(1, max_col + 1) if col_px[c - 1] > 0]

    # ── 3.5 Compute per-row heights ────────────────────────────────────────────
    # Rows that contain any Khmer text get ROW_H_KHMER, others get ROW_H
    ROW_H_KHMER = 34 * SCALE
    ROW_H = 22 * SCALE
    row_heights = []
    for r in range(1, max_row + 1):
        has_khmer_row = any(
            _has_khmer(grid[r][c]) for c in range(1, max_col + 1)
        )
        row_heights.append(ROW_H_KHMER if has_khmer_row else ROW_H)

    # ── 4. Build canvas ────────────────────────────────────────────────────────
    total_w = sum(col_px[c - 1] for c in render_cols) + 1
    total_h = sum(row_heights) + 1
    img  = Image.new('RGB', (total_w, total_h), BG_WHITE)
    draw = ImageDraw.Draw(img)

    # ── 4.5 Pre-calculate merged cell pixel dimensions ───────────────
    skip = set()
    merged_rects = {}
    for mc in ws.merged_cells.ranges:
        mw = 0
        for cc in range(mc.min_col, mc.max_col + 1):
            if cc - 1 < len(col_px) and col_px[cc - 1] > 0:
                mw += col_px[cc - 1]
        mh = 0
        for rr in range(mc.min_row, mc.max_row + 1):
            if rr - 1 < len(row_heights):
                mh += row_heights[rr - 1]

        for cc in range(mc.min_col, mc.max_col + 1):
            if (mc.min_row, cc) != (mc.min_row, mc.min_col):
                skip.add((mc.min_row, cc))
        for r_m in range(mc.min_row + 1, mc.max_row + 1):
            for c_m in range(mc.min_col, mc.max_col + 1):
                skip.add((r_m, c_m))
        merged_rects[(mc.min_row, mc.min_col)] = (mw, mh)

    # ── 5. Draw ────────────────────────────────────────────────────────────────
    y = 0
    for r in range(1, max_row + 1):
        rh = row_heights[r - 1]   # per-row height
        x = 0
        for c in render_cols:
            cw   = col_px[c - 1]
            if (r, c) in skip:
                x += cw
                continue

            mw_mh = merged_rects.get((r, c))
            if mw_mh:
                mw, mh = mw_mh
            else:
                mw, mh = cw, rh
            text = grid[r][c]

            try:
                cell  = ws.cell(r, c)
                bg    = _cell_bg(cell) or BG_WHITE
                fg    = _cell_fg(cell)
                bold  = _cell_bold(cell)
                align = _cell_align(cell)
            except Exception:
                bg, fg, bold, align = BG_WHITE, TEXT_DEF, False, 'left'

            draw.rectangle([x, y, x + mw, y + mh], fill=bg)

            if text:
                f = _get_font(text, FONT_SIZE, bold)
                try:
                    bb = draw.textbbox((0, 0), text, font=f)
                    tw, th = bb[2] - bb[0], bb[3] - bb[1]
                except Exception:
                    tw = len(text) * (FONT_SIZE - 2)
                    th = FONT_SIZE

                ty = y + (mh - th) // 2
                if align == 'center':
                    tx = x + (mw - tw) // 2
                elif align == 'right':
                    tx = x + mw - tw - PAD_X
                else:
                    tx = x + PAD_X

                draw.text((tx, ty), text, font=f, fill=fg)

            draw.rectangle([x, y, x + mw, y + mh], outline=BORDER_COL, width=1 * SCALE)
            x += cw
        y += rh

    # ── Aspect ratio safety check for Telegram ──────────────────────────────
    # Telegram rejects photos with aspect ratios > 20 (Photo_invalid_dimensions).
    # We enforce a safe ratio of 18 by adding white background padding if needed.
    max_ratio = 18.0
    w, h = img.size
    new_w, new_h = w, h
    if h > 0 and w / h > max_ratio:
        new_h = int(w / max_ratio)
    elif w > 0 and h / w > max_ratio:
        new_w = int(h / max_ratio)

    if (new_w, new_h) != (w, h):
        padded_img = Image.new('RGB', (new_w, new_h), BG_WHITE)
        padded_img.paste(img, (0, 0))
        img = padded_img

    buf = io.BytesIO()
    img.save(buf, format='PNG', optimize=True)
    buf.seek(0)
    return buf
