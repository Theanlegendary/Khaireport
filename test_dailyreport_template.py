"""
Test script to verify the new September template is being used
"""
import os
import re

# Simulate the template selection logic
template_dir = os.path.dirname(os.path.abspath(__file__))
template_files = [f for f in os.listdir(template_dir) if f.startswith("0.Master Daily Report") and f.endswith(".xlsx")]

preferred_template = "0.Master Daily Report - new - Sept_New.xlsx"

print(f"Looking for templates in: {template_dir}")
print(f"\nFound template files:")
for f in template_files:
    print(f"  - {f}")

print(f"\nPreferred template: {preferred_template}")

if os.path.exists(os.path.join(template_dir, preferred_template)):
    template_name = preferred_template
    template_path = os.path.join(template_dir, preferred_template)
    print(f"✅ Using preferred template: {template_name}")
elif template_files:
    # Sort by version number
    template_files.sort(key=lambda x: int(re.search(r'\d+', x).group()) if re.search(r'\d+', x) else 0, reverse=True)
    template_path = os.path.join(template_dir, template_files[0])
    template_name = template_files[0]
    print(f"⚠️ Preferred template not found, using: {template_name}")
else:
    template_name = "0.Master Daily Report - new - Aug.xlsx"
    template_path = os.path.join(template_dir, template_name)
    print(f"❌ No templates found, would use default: {template_name}")

print(f"\nFinal template path: {template_path}")
print(f"Template exists: {os.path.exists(template_path)}")

if os.path.exists(template_path):
    file_size = os.path.getsize(template_path)
    print(f"Template file size: {file_size:,} bytes ({file_size / 1024:.2f} KB)")
