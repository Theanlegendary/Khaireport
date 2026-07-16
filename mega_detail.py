"""mega_detail.py - Detail Excel for MEGA/HUB/DVC orders (Khmer)."""
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

STATUS_KM = {
    "110": "\u1794\u17b6\u1793\u1794\u1789\u17d2\u1787\u17b6\u1780\u17cb",
    "120": "\u1780\u17c6\u1796\u17bb\u1784\u1794\u17d2\u179a\u1798\u17bc\u179b",
    "200": "\u1794\u17b6\u1793\u1791\u1791\u17bd\u179b",
    "210": "\u1780\u17c6\u1796\u17bb\u1784\u178a\u17b9\u1780\u1787\u1789\u17d2\u1787\u17bc\u1793",
    "300": "\u1780\u17c6\u1796\u17bb\u1784\u178a\u17b9\u1780\u1787\u1789\u17d2\u1787\u17bc\u1793",
    "302": "\u1794\u17c2\u1784\u1785\u17c2\u1780\u1787\u17c4\u1782\u1787\u17d0\u1799",
    "306": "\u178a\u179b\u17cb\u1798\u1787\u17d2\u178c\u1798\u178e\u17d2\u178c\u179b",
    "309": "\u179f\u17d2\u1780\u17c1\u1793\u1794\u17bb\u1784\u1787\u17c4\u1782\u1787\u17d0\u1799",
    "310": "\u1794\u17b6\u1793\u1791\u1791\u17bd\u179b\u17a2\u17b8\u179c\u17c9\u17b6\u1793\u17cb",
    "311": "\u1794\u17c2\u1784\u1785\u17c2\u1780\u178a\u17b9\u1780\u1794\u1793\u17d2\u178f",
    "400": "\u179a\u1784\u17cb\u1785\u17b6\u17c6\u1794\u17d2\u179a\u1782\u179b\u17cb",
    "401": "\u1780\u17c6\u1796\u17bb\u1784\u178a\u17b9\u1780\u1787\u17bc\u1793",
    "402": "\u1780\u17c6\u1796\u17bb\u1784\u178a\u17b9\u1780\u1787\u17bc\u1793",
    "420": "\u179a\u1784\u17cb\u1785\u17b6\u17c6\u1799\u1780\u1793\u17c5\u17a0\u17b6\u1784",
    "430": "\u178a\u17b9\u1780\u1798\u17d2\u178f\u1784\u1791\u17c0\u178f",
    "460": "\u178a\u17b9\u1780\u1798\u17d2\u178f\u1784\u1791\u17c0\u178f",
    "472": "\u179a\u1780\u17d2\u179f\u17b6\u1791\u17bb\u1780",
    "480": "\u1780\u17c2\u17a2\u17b6\u179f\u1799\u178a\u17d2\u178b\u17b6\u1793",
    "500": "\u1795\u17d2\u1789\u17be\u178f\u17d2\u179a\u17a1\u1794\u17cb",
    "520": "\u1794\u17b6\u1793\u178f\u17d2\u179a\u17a1\u1794\u17cb",
    "540": "\u178f\u17d2\u179a\u17a1\u1794\u17cb\u1794\u179a\u17b6\u1787\u17d0\u1799",
}


# Column indices in source Excel (0-based)
CI_CREATED = 1
CI_ORDER = 2
CI_SENDER = 3
CI_RECEIVER = 4
CI_PO = 15
CI_STATUS = 23
CI_ACTION_PO_306 = 35


def _sc(val):
    if val is None:
        return ""
    s = str(val).strip()
    if " - " in s:
        s = s.split(" - ", 1)[0]
    return s.split()[0].strip() if s else ""


def build_mega_detail(source_path, out_path, cfg):
    """Build detail Excel for orders at MEGA/HUB/DVC post offices."""
    from datetime import datetime as _dt

    exclude_statuses = {"410", "201", "520", "99", "100", "-99"}
    excl_test = cfg.get("pivot", {}).get("exclude_test", False)
    test_kw = cfg.get("pivot", {}).get("test_keywords", ["test"])

    wb_src = openpyxl.load_workbook(source_path, data_only=True)
    ws_src = wb_src.active
    all_rows = list(ws_src.iter_rows(values_only=True))
    if not all_rows:
        return
    data_rows = all_rows[1:]

    today = _dt.now().date()

    filtered = []
    for row in data_rows:
        if not row or len(row) <= CI_STATUS:
            continue
        if row[CI_ORDER] is None or str(row[CI_ORDER]).strip() == "":
            continue
        sc = _sc(row[CI_STATUS])
        if sc in exclude_statuses:
            continue
        if excl_test:
            blob = " ".join(str(row[c] or "") for c in (CI_SENDER, CI_RECEIVER)).lower()
            if any(k.lower() in blob for k in test_kw):
                continue
        po = str(row[CI_PO] or "").strip().upper()
        if not ("MEGA" in po or "HUB" in po or "DVC" in po):
            continue
        filtered.append(row)

    filtered.sort(key=lambda r: (str(r[CI_PO] or ""), str(r[CI_ORDER] or "")))

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "MEGA Detail"

    km_hdr = [
        "\u1794\u17d2\u179a\u17c3\u179f\u178e\u17b7\u1799\u17cd\u1794\u1785\u17d2\u1785\u17bb\u1794\u17d2\u1794\u1793\u17d2\u1793",
        "\u179b\u17c1\u1781\u1794\u17bb\u1784",
        "\u179f\u17d2\u1790\u17b6\u1793\u1797\u17b6\u1796",
        "\u1780\u17b6\u179a\u1796\u17b7\u1796\u178e\u17cc\u1793\u17b6",
        "\u1790\u17d2\u1784\u17c3\u1794\u1784\u17d2\u1780\u17be\u178f",
        "\u17a2\u17d2\u1793\u1780\u1795\u17d2\u1789\u17be",
        "\u17a2\u17d2\u1793\u1780\u1791\u1791\u17bd\u179b",
        "\u1785\u17c6\u1793\u17bd\u1793\u1790\u17d2\u1784\u17c3",
    ]

    hfill = PatternFill("solid", fgColor="1F4E78")
    hfont = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    thin = Side(style="thin", color="BFBFBF")
    bdr = Border(left=thin, right=thin, top=thin, bottom=thin)
    ctr = Alignment(horizontal="center", vertical="center")
    urgent_fill = PatternFill("solid", fgColor="FFEBEB")

    for ci, h in enumerate(km_hdr, 1):
        c = ws.cell(1, ci, h)
        c.fill = hfill
        c.font = hfont
        c.alignment = ctr
        c.border = bdr

    def _parse_created(val):
        if val is None:
            return None
        if isinstance(val, _dt):
            return val.date()
        s = str(val).strip().split(" ")[0]
        for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%m/%d/%Y", "%d-%m-%Y"):
            try:
                return _dt.strptime(s, fmt).date()
            except ValueError:
                continue
        return None

    dfont = Font(name="Segoe UI", size=10)
    urgent_count = 0

    for ri, row in enumerate(filtered, 2):
        po = str(row[CI_PO] or "").strip()
        oid = str(row[CI_ORDER] or "").strip()
        sc = _sc(row[CI_STATUS])
        km = STATUS_KM.get(sc, sc)
        created = str(row[CI_CREATED] or "").strip()
        sender = str(row[CI_SENDER] or "").strip()
        receiver = str(row[CI_RECEIVER] or "").strip()

        # Calculate days
        created_date = _parse_created(row[CI_CREATED])
        if created_date:
            age = (today - created_date).days
        else:
            age = 0
        is_urgent = age > 1

        if is_urgent:
            urgent_count += 1

        vals = [po, oid, sc, km, created, sender, receiver, age]
        for ci, v in enumerate(vals, 1):
            c = ws.cell(ri, ci, v)
            c.font = dfont
            c.border = bdr
            c.alignment = ctr
            if is_urgent:
                c.fill = urgent_fill

    last_r = len(filtered) + 1
    if last_r > 1:
        ws.auto_filter.ref = f"A1:H{last_r}"

    for ci, w in enumerate([22, 16, 10, 28, 18, 30, 30, 12], 1):
        ws.column_dimensions[get_column_letter(ci)].width = w

    wb.save(out_path)
    return len(filtered), urgent_count
