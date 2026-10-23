"""Local signing edge: expiring source PDFs and one bounded callback route only."""
import json
import math
import re
import time
import urllib.error
import urllib.request
from pathlib import Path
from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import FileResponse, Response
from starlette.concurrency import run_in_threadpool


def create_pdf_gateway(storage):
    root = Path(storage).resolve()
    app = FastAPI(openapi_url=None, docs_url=None, redoc_url=None)

    @app.post('/esign/callback')
    async def callback(request: Request):
        if request.url.query:
            raise HTTPException(404, 'Notification introuvable.')
        body = bytearray()
        async for chunk in request.stream():
            body.extend(chunk)
            if len(body) > 32 * 1024:
                raise HTTPException(413, 'Notification trop volumineuse.')

        def forward():
            # Fixed loopback destination, no client headers/URLs or authentication
            # forwarded. The private API validates the per-document capability.
            upstream = urllib.request.Request('http://127.0.0.1:8013/esign/callback',
                data=bytes(body), headers={'Content-Type': 'application/json'}, method='POST')
            class NoRedirect(urllib.request.HTTPRedirectHandler):
                def redirect_request(self, *args, **kwargs):
                    return None
            try:
                with urllib.request.build_opener(urllib.request.ProxyHandler({}), NoRedirect()).open(upstream, timeout=10) as response:
                    status = response.status
            except urllib.error.HTTPError as error:
                status = error.code if error.code in {400, 404, 409, 413} else 503
                error.close()
            except (OSError, urllib.error.URLError):
                status = 503
            return Response(status_code=status, headers={'Cache-Control': 'no-store'})
        return await run_in_threadpool(forward)

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
