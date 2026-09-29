import sys, os
import fitz

pdf_path = sys.argv[1]
out_dir = sys.argv[2]
os.makedirs(out_dir, exist_ok=True)
doc = fitz.open(pdf_path)
paths=[]
for i, page in enumerate(doc):
    pix = page.get_pixmap(matrix=fitz.Matrix(2,2), alpha=False)
    p = os.path.join(out_dir, f'page-{i+1:04d}.png')
    pix.save(p)
    paths.append(p)
print('\n'.join(paths))
