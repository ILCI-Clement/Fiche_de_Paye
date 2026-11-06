"""Static paginated PDF from the same template cells used by register XLSX exports."""
from datetime import date, datetime
from io import BytesIO
from xml.sax.saxutils import escape

from openpyxl import load_workbook
from reportlab.lib import colors
from reportlab.lib.pagesizes import A3, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import LongTable, Paragraph, SimpleDocTemplate, Spacer, TableStyle

from .personnel_export import MAX_EXPORT_ROWS


HEADERS = ['N°', 'Nom et prénoms', 'Nationalité', 'Sexe', 'Adresse', 'N° Sécurité sociale',
           'Naissance', 'Emploi', 'Qualification', 'Entrée', 'Sortie\neffective',
           'Titre de travail', 'N° du titre', 'Contrat', 'Entreprise temporaire',
           'Apprenti', 'Jeune :\nqualification', 'Jeune :\nadaptation']


def personnel_excel_to_pdf(content, *, snapshot_at=None):
    workbook = load_workbook(BytesIO(content), data_only=True, read_only=True, keep_links=False)
    try:
        sheet = workbook.active
        if sheet.max_row > MAX_EXPORT_ROWS + 7:
            raise ValueError('Export row limit exceeded')
        rows = list(sheet.iter_rows(min_row=8, max_col=18, values_only=True))
    finally:
        workbook.close()
    output = BytesIO()
    size = landscape(A3)
    document = SimpleDocTemplate(output, pagesize=size, leftMargin=24, rightMargin=24,
                                 topMargin=28, bottomMargin=30, title='ILCI - Utilisateurs',
                                 author='ILCI - RH', invariant=1)
    body = ParagraphStyle('Cell', fontName='Helvetica', fontSize=7, leading=9)
    heading = ParagraphStyle('Head', parent=body, fontName='Helvetica-Bold', fontSize=6.5, leading=8, textColor=colors.white)
    title = ParagraphStyle('Title', fontName='Helvetica-Bold', fontSize=16, leading=20)
    def cell(value, style=body):
        if isinstance(value, (date, datetime)):
            value = value.strftime('%d/%m/%Y')
        return Paragraph(escape(str(value if value is not None else '')).replace('\n', '<br/>'), style)
    story = [Paragraph('ILCI - Utilisateurs', title), Spacer(1, 8)]
    if snapshot_at:
        story += [Paragraph(f"État historique au {snapshot_at.strftime('%d/%m/%Y %H:%M')} UTC", body), Spacer(1, 8)]
    weights = [24, 100, 60, 30, 110, 80, 60, 80, 80, 60, 60, 80, 80, 75, 100, 40, 55, 55]
    widths = [(size[0] - 48) * weight / sum(weights) for weight in weights]
    if rows:
        table = LongTable([[cell(label, heading) for label in HEADERS]] + [[cell({'Homme': 'H', 'Femme': 'F'}.get(value, value) if index == 3 else value) for index, value in enumerate(row)] for row in rows],
                          colWidths=widths, repeatRows=1, splitByRow=1, splitInRow=1)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2859A0')),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F1F4F8')]),
            ('GRID', (0, 0), (-1, -1), .35, colors.HexColor('#CDD4DF')),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING', (0, 0), (-1, -1), 4), ('RIGHTPADDING', (0, 0), (-1, -1), 4),
            ('TOPPADDING', (0, 0), (-1, -1), 6), ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ]))
        story.append(table)
    else:
        story.append(Paragraph('Aucune personne dans cette archive.', body))
    def footer(canvas, _document):
        canvas.saveState()
        canvas.setFont('Helvetica', 8)
        canvas.setFillColor(colors.HexColor('#596579'))
        canvas.drawString(24, 15, 'ILCI - RH | Document confidentiel')
        canvas.drawRightString(size[0] - 24, 15, f'Page {canvas.getPageNumber()}')
        canvas.restoreState()
    document.build(story, onFirstPage=footer, onLaterPages=footer)
    return output.getvalue()
