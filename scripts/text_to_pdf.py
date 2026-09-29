import sys
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.enums import TA_LEFT
from xml.sax.saxutils import escape

out, title = sys.argv[1], sys.argv[2]
text = open(sys.argv[3], 'r', encoding='utf-8', errors='replace').read()
styles = getSampleStyleSheet()
style = styles['BodyText']; style.fontName='Helvetica'; style.fontSize=10; style.leading=14; style.alignment=TA_LEFT
story = [Paragraph(escape(title or 'Dokumen'), styles['Title']), Spacer(1, 12)]
for line in text.replace('\r\n','\n').replace('\r','\n').split('\n'):
    story.append(Paragraph(escape(line) if line else '&nbsp;', style))
    story.append(Spacer(1, 4))
SimpleDocTemplate(out, pagesize=A4, rightMargin=36,leftMargin=36,topMargin=36,bottomMargin=36).build(story)
print(out)
