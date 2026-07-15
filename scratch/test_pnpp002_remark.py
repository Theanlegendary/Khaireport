import json
import sys
import os
sys.path.insert(0, os.path.abspath("c:/Users/DELL/Desktop/daily_push"))
import generate_report

with open("config.json") as f:
    cfg = json.load(f)

res = generate_report.generate_reports_from_data(
    "test_detail.xlsx",
    "post_office_lookup.csv",
    "test_out",
    return_metadata=True,
    target_handles=["PNPP002"]
)

for hr in res["handle_results"]:
    print("New Remark:", hr["remark"])
    print("New Counts:", hr["handle_counts"])
