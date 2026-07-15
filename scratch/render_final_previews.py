import os
import sys

# Ensure paths
sys.path.insert(0, os.path.abspath("c:/Users/DELL/Desktop/daily_push"))

import excel_to_image

out_dir = "c:/Users/DELL/Desktop/daily_push/test_out"
artifact_dir = "c:/Users/DELL/.gemini/antigravity/brain/b8105454-e4a4-4a86-b723-6ef7dcbedc27"

# Get all files sorted by modification time descending
files_with_time = [
    (os.path.join(out_dir, f), os.path.getmtime(os.path.join(out_dir, f)))
    for f in os.listdir(out_dir)
    if os.path.isfile(os.path.join(out_dir, f))
]
# Sort by mtime desc
files_with_time.sort(key=lambda x: x[1], reverse=True)
latest_files = [x[0] for x in files_with_time]

# Find latest Delivery and Pending files for BANP001
delivery_file = next((f for f in latest_files if "BANP001" in f and "Delivery" in f), None)
pending_file = next((f for f in latest_files if "BANP001" in f and "Pending" in f), None)

# Render files
if delivery_file:
    print("Rendering final Delivery...")
    img = excel_to_image.excel_to_image(delivery_file)
    with open(os.path.join(artifact_dir, "final_delivery.png"), "wb") as f:
        f.write(img.getvalue())
        
if pending_file:
    print("Rendering final Pending...")
    img = excel_to_image.excel_to_image(pending_file)
    with open(os.path.join(artifact_dir, "final_pending.png"), "wb") as f:
        f.write(img.getvalue())

print("SUCCESS")
