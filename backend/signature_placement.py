"""Resolve template anchors and map PDF points to ClawShow's A4 field space.

ClawShow's current field renderer scales x/width by page_width/595 and
y/height by page_height/842, including fields obtained from anchors.
Sending explicit normalized coordinates avoids double scaling landscape anchors.
"""
from io import BytesIO
from pypdf import PdfReader
from fastapi import HTTPException

PROVIDER_WIDTH = 595.0
PROVIDER_HEIGHT = 842.0


def place_signature_fields(pdf_bytes, anchors):
    try:
        reader = PdfReader(BytesIO(pdf_bytes))
        positions = {field['name']: [] for field in anchors}
        for page_number, page in enumerate(reader.pages, 1):
            if page.rotation or list(page.cropbox) != list(page.mediabox):
                raise HTTPException(400, 'La page PDF tournée ou recadrée doit être régénérée avant signature.')
            width, height = float(page.mediabox.width), float(page.mediabox.height)
            left, bottom = float(page.mediabox.left), float(page.mediabox.bottom)
            def visit(text, cm, tm, font, font_size):
                for field in anchors:
                    # Refuse ambiguous chunks rather than guessing an offset inside text.
                    if text.strip() == field['anchor']:
                        x = tm[4] * cm[0] + tm[5] * cm[2] + cm[4] - left
                        baseline = tm[4] * cm[1] + tm[5] * cm[3] + cm[5] - bottom
                        y = height - baseline - font_size * 0.8 + field['dy']
                        x += field['dx']
                        w, h = field['width'], field['height']
                        if x < 0 or y < 0 or x+w > width or y+h > height:
                            raise HTTPException(400, 'Le cadre de signature dépasse la page PDF.')
                        positions[field['name']].append({
                            'name':field['name'], 'type':field['type'], 'party':field['party'],
                            'page':page_number,
                            'x':round(x*PROVIDER_WIDTH/width, 3),
                            'y':round(y*PROVIDER_HEIGHT/height, 3),
                            'width':round(w*PROVIDER_WIDTH/width, 3),
                            'height':round(h*PROVIDER_HEIGHT/height, 3),
                        })
            page.extract_text(visitor_text=visit)
        if any(len(found) != 1 for found in positions.values()):
            raise HTTPException(400, 'Chaque libellé de signature doit apparaître une seule fois dans le PDF. Régénérez la fiche.')
        return [positions[field['name']][0] for field in anchors]
    except HTTPException:
        raise
    except Exception as error:
        raise HTTPException(400, 'Impossible de déterminer les positions de signature dans le PDF.') from error
