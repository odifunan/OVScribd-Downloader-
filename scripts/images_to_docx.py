import sys, os
from docx import Document
from docx.shared import Inches
from docx.enum.section import WD_SECTION
from PIL import Image

out = sys.argv[1]
images = sys.argv[2:]
if not images:
    raise SystemExit('Tidak ada halaman gambar.')

doc = Document()
sec = doc.sections[0]
sec.top_margin = Inches(0)
sec.bottom_margin = Inches(0)
sec.left_margin = Inches(0)
sec.right_margin = Inches(0)

for i, p in enumerate(images):
    if i:
        sec = doc.add_section(WD_SECTION.NEW_PAGE)
        sec.top_margin = Inches(0); sec.bottom_margin = Inches(0)
        sec.left_margin = Inches(0); sec.right_margin = Inches(0)
    with Image.open(p) as im:
        w, h = im.size
    # Fit image to printable page area while preserving aspect ratio.
    max_w = 8.27 if h >= w else 11.69
    max_h = 11.69 if h >= w else 8.27
    scale = min(max_w / (w / 150), max_h / (h / 150))
    width = (w / 150) * scale
    height = (h / 150) * scale
    para = doc.add_paragraph()
    para.paragraph_format.space_after = 0
    para.paragraph_format.space_before = 0
    para.add_run().add_picture(p, width=Inches(width), height=Inches(height))

doc.save(out)
print(out)
