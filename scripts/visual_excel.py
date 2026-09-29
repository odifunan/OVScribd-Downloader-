import sys, os, json
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
from PIL import Image

# Usage: visual_excel.py output.xlsx image1 image2 ...
out = sys.argv[1]
paths = sys.argv[2:]
if not paths:
    raise SystemExit('Tidak ada gambar halaman.')

wb = Workbook()
ws = wb.active
ws.title = 'Tampilan Asli'
ws.sheet_view.showGridLines = False

# Excel column width is approximately 7 px per unit; use a wide canvas.
for c in range(1, 80):
    ws.column_dimensions[get_column_letter(c)].width = 2.1

for idx, p in enumerate(paths, 1):
    if not os.path.exists(p):
        continue
    with Image.open(p) as im:
        w, h = im.size
    img = XLImage(p)
    # Preserve the rendered page aspect ratio and use a readable page width.
    target_w = 900
    target_h = int(h * target_w / w)
    img.width = target_w
    img.height = target_h
    anchor = f'A{1 if idx == 1 else (ws.max_row + 3)}'
    ws.add_image(img, anchor)
    # Reserve approximate row space for the image.
    rows_needed = max(45, int(target_h / 18))
    start = ws[anchor].row
    for r in range(start, start + rows_needed):
        ws.row_dimensions[r].height = 18
    ws.cell(start, 1).value = f'Halaman {idx}'
    ws.cell(start, 1).font = __import__('openpyxl').styles.Font(bold=True)
    # Move next insertion point by setting a hidden marker row.
    ws.row_dimensions[start + rows_needed].height = 8

wb.save(out)
