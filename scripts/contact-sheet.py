"""Lays the slide pictures of a deck out on one sheet, for a README or a quick look.

    python scripts/contact-sheet.py <folder of slide-NN.png> <out.png> [columns] [thumbnail width]

Make the slide pictures with scripts/render-pptx.ps1 first. Needs Pillow.
"""
import glob
import os
import sys

from PIL import Image, ImageDraw

source, out = sys.argv[1], sys.argv[2]
columns = int(sys.argv[3]) if len(sys.argv) > 3 else 3
width = int(sys.argv[4]) if len(sys.argv) > 4 else 560

files = sorted(glob.glob(os.path.join(source, 'slide-*.png')))
if not files:
    sys.exit(f'no slide-*.png in {source}')
slides = [Image.open(file).convert('RGB') for file in files]
height = round(width * slides[0].height / slides[0].width)
gap = 14
rows = (len(slides) + columns - 1) // columns
sheet = Image.new('RGB', (columns * (width + gap) + gap, rows * (height + gap) + gap), (244, 245, 247))
draw = ImageDraw.Draw(sheet)
for index, slide in enumerate(slides):
    x = gap + (index % columns) * (width + gap)
    y = gap + (index // columns) * (height + gap)
    sheet.paste(slide.resize((width, height), Image.LANCZOS), (x, y))
    draw.rectangle([x - 1, y - 1, x + width, y + height], outline=(214, 217, 222))
os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
sheet.save(out, optimize=True)
print(f'{out}: {sheet.width}x{sheet.height}, {os.path.getsize(out) // 1024} KB')
