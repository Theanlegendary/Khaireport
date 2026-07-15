import os
import sys

# Ensure current dir is in sys.path
sys.path.insert(0, os.path.abspath("c:/Users/DELL/Desktop/daily_push"))
sys.path.insert(0, os.path.abspath("c:/Users/DELL/Desktop/daily_push/scratch"))

import generate_report_temp
import excel_to_image

src_xlsx = "c:/Users/DELL/Desktop/daily_push/test_detail.xlsx"
ref_path = "c:/Users/DELL/Desktop/daily_push/post_office_lookup.csv"
out_dir = "c:/Users/DELL/Desktop/daily_push/test_out_khmer"

print("Running temp generator...")
result = generate_report_temp.generate_reports_from_data(
    src_xlsx, ref_path, out_dir, return_metadata=True, mode="wide"
)

# Find first handle with files
first_hr = result["handle_results"][0]
print(f"First handle: {first_hr['handle']}")
print("Generated files:")
# Find Transit and Action Needed files
transit_file = next((f['path'] for f in first_hr["handle_files"] if "Transit" in f['path']), None)
action_file = next((f['path'] for f in first_hr["handle_files"] if "Action" in f['path']), None)

artifact_dir = "c:/Users/DELL/.gemini/antigravity/brain/b8105454-e4a4-4a86-b723-6ef7dcbedc27"

if action_file:
    print("Rendering Action Needed...")
    img_buf = excel_to_image.excel_to_image(action_file)
    out_png = os.path.join(artifact_dir, "test_action_needed.png")
    with open(out_png, "wb") as f_img:
        f_img.write(img_buf.getvalue())
    print("Successfully saved Action Needed image")

if transit_file:
    print("Rendering Transit...")
    img_buf = excel_to_image.excel_to_image(transit_file)
    out_png = os.path.join(artifact_dir, "test_transit.png")
    with open(out_png, "wb") as f_img:
        f_img.write(img_buf.getvalue())
    print("Successfully saved Transit image")

print("SUCCESS")
