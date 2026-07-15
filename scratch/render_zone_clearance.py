import os
import sys

# Add directory to sys.path to be able to import local files
sys.path.append(r"c:\Users\DELL\Desktop\daily_push")

from excel_to_image import excel_to_image

xlsx_path = r"c:\Users\DELL\Desktop\daily_push\Zone_Clearance_Report_13_07_2026.xlsx"
out_img_path = r"c:\Users\DELL\.gemini\antigravity\brain\b8105454-e4a4-4a86-b723-6ef7dcbedc27\zone_clearance_report.png"

print(f"Rendering {xlsx_path} to image using excel_to_image...")
try:
    img_data = excel_to_image(xlsx_path)
    
    # Save the bytes to file
    with open(out_img_path, "wb") as f:
        f.write(img_data.getvalue())
        
    print(f"Image successfully rendered and saved to: {out_img_path}")
except Exception as e:
    print(f"Failed to render image: {e}")
