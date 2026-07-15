"""Test if Khmer fonts can be loaded and render correctly"""
from PIL import Image, ImageDraw, ImageFont
import io

def test_font(font_path, khmer_text="ចាត់តាំងអ្នកដឹក"):
    try:
        font = ImageFont.truetype(font_path, 24)
        img = Image.new("RGB", (400, 100), (255, 255, 255))
        draw = ImageDraw.Draw(img)
        draw.text((10, 10), khmer_text, font=font, fill=(0, 0, 0))
        img.save(f"test_{font_path.split('/')[-1].replace('.ttf', '')}.png")
        print(f"✓ {font_path} — SUCCESS")
        return True
    except Exception as e:
        print(f"✗ {font_path} — FAILED: {e}")
        return False

# Test all available Khmer fonts
fonts = [
    "C:/Windows/Fonts/KhmerOSbattambang.ttf",
    "C:/Windows/Fonts/KhmerOSsiemreap.ttf",
    "C:/Windows/Fonts/KhmerOScontent.ttf",
    "C:/Windows/Fonts/KhmerOS.ttf",
    "C:/Windows/Fonts/KhmerOSsys.ttf",
]

print("Testing Khmer font rendering...\n")
for font_path in fonts:
    test_font(font_path)

print("\nCheck the test_*.png files to see if Khmer text renders correctly.")
