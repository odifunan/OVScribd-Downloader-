import sys, os
import fitz
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font
from openpyxl.worksheet.page import PageMargins

pdf_path, out_xlsx = sys.argv[1], sys.argv[2]
render_dir = sys.argv[3]
os.makedirs(render_dir, exist_ok=True)

doc = fitz.open(pdf_path)
if len(doc) == 0:
    raise SystemExit('PDF tidak memiliki halaman.')

wb = Workbook()
# Remove default sheet after creating the first real page.
wb.remove(wb.active)

for idx, page in enumerate(doc, 1):
    ws = wb.create_sheet(f'Halaman {idx}')
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = None
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = 'landscape' if page.rect.width >= page.rect.height else 'portrait'
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_margins = PageMargins(left=0, right=0, top=0, bottom=0, header=0, footer=0)
    ws.print_options.horizontalCentered = True
    ws.print_options.verticalCentered = True
    ws.sheet_properties.outlinePr.summaryBelow = True

    # 150 DPI is a good compromise between visual fidelity and workbook size.
    scale = 150 / 72
    pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
    img_path = os.path.join(render_dir, f'page-{idx:04d}.png')
    pix.save(img_path)

    img = XLImage(img_path)
    img.width = pix.width
    img.height = pix.height
    ws.add_image(img, 'A1')

    # Give Excel a printable canvas roughly matching the rendered page.
    ws.column_dimensions['A'].width = max(1, pix.width / 7.0)
    rows = max(1, int(pix.height / 18))
    for r in range(1, rows + 1):
        ws.row_dimensions[r].height = 18
    ws.print_area = f'A1:A{rows}'
    ws['A1'].font = Font(size=1, color='FFFFFF')
    ws['A1'].value = ''

wb.save(out_xlsx)
print(out_xlsx)
