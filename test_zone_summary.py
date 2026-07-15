import os
import json
import generate_report
import generate_summary

with open("config.json", encoding="utf-8") as f:
    cfg = json.load(f)

src_xlsx = "test_detail.xlsx"
ref_path = "post_office_lookup.csv"
out_dir = "test_out"

# Generate report metadata
result = generate_report.generate_reports_from_data(
    src_xlsx, ref_path, out_dir, return_metadata=True, mode="wide"
)

# Apply Zone 5 filter just like bot.py does for /total zone5
zone_key = "zone5"
zone_filter = [h.upper() for h in cfg["total_zones"][zone_key]]

# Filter handle results
filtered_results = [
    hr for hr in result["handle_results"]
    if hr["handle"] in zone_filter
]

# Recalculate overall counts
overall = {"Pickup": 0, "Delivery": 0, "Pending": 0}
for hr in filtered_results:
    for k in overall:
        overall[k] += hr["handle_counts"].get(k, 0)

# Build summary image using generate_summary
img_buf = generate_summary.build_summary_image(filtered_results, overall)

# Save image
out_png = "zone5_summary_test.png"
with open(out_png, "wb") as f_img:
    f_img.write(img_buf.getvalue())

print("Successfully generated Zone 5 summary image!")
