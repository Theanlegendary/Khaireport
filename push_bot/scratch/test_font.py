import sys
import os
from PIL import Image, ImageDraw, ImageFont

print("Python version:", sys.version)

font_dirs = [
    "C:\\Windows\\Fonts",
]

test_fonts = [
    "LeelawUI.ttf",
    "LeelaUIb.ttf",
    "KhmerOS.ttf",
    "KhmerOScontent.ttf",
    "arial.ttf"
]

for fd in font_dirs:
    for tf in test_fonts:
        path = os.path.join(fd, tf)
        exists = os.path.exists(path)
        print(f"Font: {tf} at {path} exists: {exists}")
        if exists:
            try:
                font = ImageFont.truetype(path, 12)
                print(f"  Successfully loaded via absolute path!")
            except Exception as e:
                print(f"  Failed loading absolute path: {e}")
            try:
                font2 = ImageFont.truetype(tf, 12)
                print(f"  Successfully loaded via name '{tf}'!")
            except Exception as e:
                print(f"  Failed loading name '{tf}': {e}")
