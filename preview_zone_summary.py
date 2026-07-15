"""Quick preview — generates zone summary image with new layout using fake data."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from datetime import date, datetime
import generate_summary

# ── Fake data ──────────────────────────────────────────────────────────────────
handle_results = [
    {"handle": "BANP001", "handle_counts": {"Pickup": 4, "Delivery": 2, "Pending": 5}},
    {"handle": "BATP001", "handle_counts": {"Pickup": 0, "Delivery": 3, "Pending": 8}},
    {"handle": "CHHP001", "handle_counts": {"Pickup": 2, "Delivery": 1, "Pending": 3}},
    {"handle": "PURP001", "handle_counts": {"Pickup": 1, "Delivery": 4, "Pending": 2}},
]

overall = {"Pickup": 7, "Delivery": 10, "Pending": 18}

# Fake per-date bill counts
d1, d2, d3 = date(2026, 7, 4), date(2026, 7, 5), date(2026, 7, 6)
day_date_counts = {
    "BANP001": {d1: 3, d2: 4, d3: 4},
    "BATP001": {d1: 2, d2: 3, d3: 6},
    "CHHP001": {d1: 1, d2: 2, d3: 3},
    "PURP001": {d1: 0, d2: 3, d3: 4},
}

# Urgent = overdue bills
urgent_counts = {
    "BANP001": 3,
    "BATP001": 6,
    "CHHP001": 0,
    "PURP001": 2,
}

buf = generate_summary.build_summary_image(
    handle_results,
    overall,
    zone_label="ZONE3",
    day_date_counts=day_date_counts,
    urgent_counts=urgent_counts,
    today=datetime(2026, 7, 6, 9, 30),
)

out = "zone_summary_preview.png"
with open(out, "wb") as f:
    f.write(buf.getvalue())

print(f"Saved: {out}")
