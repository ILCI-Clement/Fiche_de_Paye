"""PDF-only public edge for the local sandbox; never mounts the private API."""
import json
import math
import re
import time
from pathlib import Path
from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import FileResponse


def create_pdf_gateway(storage):
    root = Path(storage).resolve()
    app = FastAPI(openapi_url=None, docs_url=None, redoc_url=None)

    @app.api_route('/esign/documents/{token}', methods=['GET', 'HEAD'])
    def document(token: str, request: Request):
        if not re.fullmatch(r'[A-Za-z0-9_-]{43}', token) or request.url.query:
            raise HTTPException(404, 'Document introuvable.')
        metadata_path = root / f'{token}.json'
        pdf = root / f'{token}.pdf'
        try:
            if metadata_path.resolve().parent != root or pdf.resolve().parent != root:
                raise ValueError()
            metadata = json.loads(metadata_path.read_text(encoding='utf-8'))
            expires = float(metadata['expires_at'])
            if not math.isfinite(expires) or expires <= time.time() or metadata['pdf_name'] != pdf.name or not pdf.is_file():
                raise ValueError()
        except (OSError, ValueError, KeyError, TypeError):
            raise HTTPException(404, 'Document introuvable.') from None
        return FileResponse(pdf, media_type='application/pdf', filename='document.pdf',
                            headers={'Cache-Control': 'private, no-store', 'X-Content-Type-Options': 'nosniff',
                                     'Referrer-Policy': 'no-referrer'})

    return app


app = create_pdf_gateway(Path(__file__).resolve().parents[1] / '.local-test' / 'rebuild' / 'esign-documents')
