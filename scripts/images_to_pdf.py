import sys
from PIL import Image

if len(sys.argv) < 3:
    raise SystemExit('usage: images_to_pdf.py OUTPUT IMAGE...')
out = sys.argv[1]
paths = sys.argv[2:]
imgs = []
for p in paths:
    im = Image.open(p).convert('RGB')
    imgs.append(im)
if not imgs:
    raise SystemExit('no images')
imgs[0].save(out, 'PDF', resolution=150.0, save_all=True, append_images=imgs[1:])
