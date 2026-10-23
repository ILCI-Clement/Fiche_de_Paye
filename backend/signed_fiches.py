"""Receive ClawShow completions, archive original PDFs, expose authorized downloads."""
import hashlib
import io
import json
import logging
import os
from pathlib import Path
import re
import tempfile
import threading
import time
import urllib.request
from urllib.parse import quote, urlsplit
from zipfile import ZIP_DEFLATED, ZipFile

from fastapi import Depends, HTTPException, Query, Request
from fastapi.responses import Response
from starlette.concurrency import run_in_threadpool
from pypdf import PdfReader

from document_downloads import MAX_ARCHIVE_BYTES, safe_archive_component
if __package__:
    from .signed_fiche_store import SignedFicheStore, ensure_signed_fiche_schema, event
    from .role_policy import checked_resource_actor, is_admin
else:
    from signed_fiche_store import SignedFicheStore, ensure_signed_fiche_schema, event
    from role_policy import checked_resource_actor, is_admin

MAX_SIGNED_PDF_BYTES = 25 * 1024 * 1024
PRIVATE_HEADERS = {'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff'}


def callback_url():
    value = os.getenv('PRESENCE_ESIGN_CALLBACK_URL', '').strip()
    if not value:
        return None
    try:
        url = urlsplit(value)
        if (url.scheme != 'https' or not url.hostname or url.hostname in {'localhost', '127.0.0.1', '::1'}
                or url.username or url.password or url.query or url.fragment or url.port not in {None, 443}
                or not url.path.endswith('/esign/callback')):
            raise ValueError()
    except ValueError:
        raise HTTPException(503, 'L’adresse de retour des signatures doit être configurée.') from None
    return value


def validate_signed_url(value, document_id):
    # Never fetch arbitrary callback URLs, redirects or a preview/original PDF.
    if not isinstance(value, str) or len(value) > 4096:
        raise HTTPException(400, 'Adresse du PDF signé invalide.')
    try:
        url = urlsplit(value)
        if (url.scheme != 'https' or url.hostname != 'esign.clawshow.ai' or url.port not in {None, 443}
                or url.username or url.password or url.fragment
                or url.path != f'/esign/{document_id}/signed.pdf'):
            raise ValueError()
    except ValueError:
        raise HTTPException(400, 'Adresse du PDF signé non autorisée.') from None
    return value


class NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        return None


def fetch_signed_pdf(row):
    url = validate_signed_url(row['signed_pdf_url'], row['document_id'])
    headers = {'Accept': 'application/pdf'}
    key = os.getenv('CLAWSHOW_ESIGN_API_KEY', '')
    # Credential sent only to the exact official host; redirects are forbidden.
    if key:
        headers['Authorization'] = f'Bearer {key}'
    started = time.monotonic()
    with urllib.request.build_opener(NoRedirect()).open(urllib.request.Request(url, headers=headers), timeout=10) as response:
        if response.status != 200:
            raise ValueError('PDF unavailable')
        length = response.headers.get('Content-Length')
        if length and int(length) > MAX_SIGNED_PDF_BYTES:
            raise ValueError('PDF too large')
        output = io.BytesIO()
        while chunk := response.read(64 * 1024):
            if output.tell() + len(chunk) > MAX_SIGNED_PDF_BYTES or time.monotonic() - started > 60:
                raise ValueError('PDF limit exceeded')
            output.write(chunk)
    data = output.getvalue()
    if not data.startswith(b'%PDF-') or b'%%EOF' not in data[-4096:]:
        raise ValueError('Not a final PDF')
    pdf = PdfReader(io.BytesIO(data))
    if pdf.is_encrypted or not 1 <= len(pdf.pages) <= 300:
        raise ValueError('Invalid PDF')
    # Validation only. Rewriting a signed PDF would invalidate its signatures.
    return data


class SignedFiches:
    def __init__(self, connect, root):
        self.store = SignedFicheStore(connect)
        self.root = Path(root).resolve()
        self.stop = threading.Event()
        self.wake = threading.Event()
        self.thread = None

    def prepare(self, **values):
        url = callback_url()
        if not url:
            return None
        archive_id, reference = self.store.prepare(**values)
        return {'id': archive_id, 'reference': reference, 'callback_url': url}

    def path(self, row):
        if (not re.fullmatch(r'[0-9a-f]{32}', row['id']) or not 2000 <= row['fiche_year'] <= 2100
                or not 1 <= row['fiche_month'] <= 12):
            raise ValueError('Invalid archive metadata')
        path = (self.root / str(row['fiche_year']) / f"{row['fiche_month']:02d}" / f"{row['id']}.pdf").resolve()
        if not path.is_relative_to(self.root):
            raise ValueError('Invalid archive path')
        return path

    def save(self, row, data):
        path = self.path(row)
        for directory in (self.root, path.parent.parent, path.parent):
            directory.mkdir(parents=True, exist_ok=True, mode=0o700)
        digest = hashlib.sha256(data).hexdigest()
        temporary_name = None
        try:
            with tempfile.NamedTemporaryFile(dir=path.parent, suffix='.part', delete=False) as temporary:
                temporary_name = Path(temporary.name)
                temporary.write(data)
                temporary.flush()
                os.fsync(temporary.fileno())
            os.chmod(temporary_name, 0o600)
            try:
                # Atomic publication without ever replacing an existing signed file.
                os.link(temporary_name, path)
            except FileExistsError:
                if hashlib.sha256(path.read_bytes()).hexdigest() != digest:
                    raise ValueError('Archived PDF differs') from None
            if os.name != 'nt':
                descriptor = os.open(path.parent, os.O_RDONLY | os.O_DIRECTORY)
                try:
                    os.fsync(descriptor)
                finally:
                    os.close(descriptor)
        finally:
            if temporary_name:
                temporary_name.unlink(missing_ok=True)
        self.store.downloaded(row, digest, len(data))

    def read(self, row):
        if row['status'] != 'ready':
            raise HTTPException(409, 'Le PDF signé n’est pas encore archivé.')
        try:
            path = self.path(row)
            if path.stat().st_size > MAX_SIGNED_PDF_BYTES:
                raise ValueError()
            data = path.read_bytes()
            if len(data) != row['file_bytes'] or hashlib.sha256(data).hexdigest() != row['sha256']:
                raise ValueError()
            return data
        except (OSError, ValueError):
            raise HTTPException(503, 'Le fichier archivé doit être vérifié. Aucun fichier partiel n’a été fourni.') from None

    def run(self):
        while not self.stop.is_set():
            self.wake.clear()
            try:
                while not self.stop.is_set():
                    row = self.store.claim()
                    if not row:
                        break
                    try:
                        self.save(row, fetch_signed_pdf(row))
                    except Exception:
                        # No raw provider error, signed URL, reference or personal data in logs.
                        self.store.failed(row)
            except Exception:
                logging.getLogger(__name__).warning('Signed fiche archive worker deferred; retry scheduled.')
            self.wake.wait(15)


def filename(row):
    name = safe_archive_component(row['employee_name'])
    return f"{name}_{row['fiche_year']}-{row['fiche_month']:02d}_{row['id']}_signe.pdf"


def file_response(data, name, media_type='application/pdf'):
    return Response(data, media_type=media_type, headers={**PRIVATE_HEADERS,
                    'Content-Disposition': f"attachment; filename*=UTF-8''{quote(name)}"})


def register_signed_fiche_routes(app, *, get_db_connection, get_current_user, load_user, root):
    service = SignedFiches(get_db_connection, root)
    app.state.signed_fiches = service

    @app.on_event('startup')
    def startup():
        with service.store.transaction() as cursor:
            ensure_signed_fiche_schema(cursor)
        callback_url()  # Reject an invalid configured address before any real send.
        if os.getenv('PRESENCE_SIGNED_FICHE_WORKER', '1') == '1':
            service.thread = threading.Thread(target=service.run, daemon=True, name='signed-fiche-archive')
            service.thread.start()

    @app.on_event('shutdown')
    def shutdown():
        service.stop.set()
        service.wake.set()
        if service.thread:
            service.thread.join(timeout=2)

    def current_actor(cursor, actor):
        return checked_resource_actor(cursor, actor, load_user)

    def allowed(cursor, actor, row):
        target = load_user(cursor, row['employee_username'])
        if not (is_admin(actor) or actor['username'] == row['employee_username'] or
                ('Responsable' in actor['role_tags'] and target and target.get('manager_username') == actor['username'])):
            raise HTTPException(403, 'Ce document est réservé à la personne, son responsable direct et un administrateur.')

    def get_row(cursor, archive_id):
        cursor.execute('SELECT * FROM signed_fiche_documents WHERE id=%s', (archive_id,))
        row = cursor.fetchone()
        if not row:
            raise HTTPException(404, 'Document signé introuvable.')
        return row

    @app.post('/esign/callback', include_in_schema=False)
    async def callback(request: Request):
        raw = bytearray()
        async for chunk in request.stream():
            raw.extend(chunk)
            if len(raw) > 32 * 1024:
                raise HTTPException(413, 'Notification trop volumineuse.')
        try:
            payload = json.loads(raw)
            if not isinstance(payload, dict):
                raise ValueError()
        except (ValueError, UnicodeError):
            raise HTTPException(400, 'Notification invalide.') from None
        # Hash lookup of the unpredictable per-request reference authenticates this
        # body. ClawShow currently documents no HMAC header. Never log this payload.
        await run_in_threadpool(service.store.receive, payload, validate_signed_url)
        service.wake.set()
        return Response(status_code=204, headers=PRIVATE_HEADERS)

    @app.get('/signed-fiches')
    def listing(response: Response, year: int = Query(ge=2000, le=2100), month: int = Query(ge=1, le=12),
                mode: str = Query('production', pattern='^(production|sandbox-only)$'),
                username: str | None = Query(None, max_length=255),
                page: int = Query(1, ge=1), actor=Depends(get_current_user)):
        response.headers.update(PRIVATE_HEADERS)
        with service.store.transaction() as cursor:
            actor = current_actor(cursor, actor)
            where = 'fiche_year=%s AND fiche_month=%s AND mode=%s'
            args = [year, month, mode]
            if username is not None:
                allowed(cursor, actor, {'employee_username': username})
                where += ' AND employee_username=%s'
                args.append(username)
            if not is_admin(actor):
                where += ''' AND (employee_username=%s OR employee_username IN
                    (SELECT username FROM users WHERE manager_username=%s AND %s=1))'''
                args += [actor['username'], actor['username'], int('Responsable' in actor['role_tags'])]
            cursor.execute(f'SELECT COUNT(*) AS total FROM signed_fiche_documents WHERE {where}', args)
            total = cursor.fetchone()['total']
            page = min(page, max(1, (total + 19) // 20))
            cursor.execute(f'''SELECT id,document_id,employee_name,fiche_month,fiche_year,mode,status,
                created_at,completed_at,archived_at,file_bytes,last_error FROM signed_fiche_documents
                WHERE {where} ORDER BY created_at DESC,id LIMIT 20 OFFSET %s''', [*args, (page - 1) * 20])
            items = cursor.fetchall()
        return {'items': items, 'total': total, 'page': page, 'can_export': is_admin(actor),
                'automatic_archive': bool(callback_url())}

    @app.get('/signed-fiches/{archive_id}/file')
    def download(archive_id: str, actor=Depends(get_current_user)):
        with service.store.transaction() as cursor:
            actor = current_actor(cursor, actor)
            row = get_row(cursor, archive_id)
            allowed(cursor, actor, row)
            data = service.read(row)
            event(cursor, archive_id, 'downloaded', actor['username'])
        return file_response(data, filename(row))

    @app.post('/signed-fiches/{archive_id}/retry')
    def retry(archive_id: str, actor=Depends(get_current_user)):
        with service.store.transaction() as cursor:
            actor = current_actor(cursor, actor)
            if not is_admin(actor):
                raise HTTPException(403, 'Action réservée aux administrateurs.')
            row = get_row(cursor, archive_id)
            if row['status'] != 'completed':
                raise HTTPException(409, 'Seul le téléchargement d’un document terminé peut être relancé.')
            cursor.execute('UPDATE signed_fiche_documents SET next_attempt_at=UTC_TIMESTAMP(6) WHERE id=%s', (archive_id,))
            event(cursor, archive_id, 'download_retry', actor['username'])
        service.wake.set()
        return {'status': 'queued'}

    @app.get('/signed-fiches/export.zip')
    def export(year: int = Query(ge=2000, le=2100), month: int = Query(ge=1, le=12),
               mode: str = Query('production', pattern='^(production|sandbox-only)$'), actor=Depends(get_current_user)):
        with service.store.transaction() as cursor:
            actor = current_actor(cursor, actor)
            if not is_admin(actor):
                raise HTTPException(403, 'Export mensuel réservé aux administrateurs.')
            cursor.execute('''SELECT * FROM signed_fiche_documents
                WHERE fiche_year=%s AND fiche_month=%s AND mode=%s AND status IN ('completed','ready')
                ORDER BY employee_name,id LIMIT 5001''', (year, month, mode))
            rows = cursor.fetchall()
            if not rows:
                raise HTTPException(404, 'Aucun PDF signé archivé pour ce mois.')
            if any(row['status'] != 'ready' for row in rows):
                raise HTTPException(409, 'Des documents signés attendent leur archivage. Réessayez après leur récupération.')
            if len(rows) > 5000 or sum(row['file_bytes'] or 0 for row in rows) > MAX_ARCHIVE_BYTES:
                raise HTTPException(413, 'L’export dépasse 5 000 fichiers ou 100 Mo. Téléchargez les fichiers séparément.')
            output = io.BytesIO()
            total = 0
            with ZipFile(output, 'w', compression=ZIP_DEFLATED) as archive:
                for row in rows:
                    data = service.read(row)
                    total += len(data)
                    if total > MAX_ARCHIVE_BYTES:
                        raise HTTPException(413, 'L’export dépasse 100 Mo. Aucun ZIP partiel fourni.')
                    archive.writestr(f"{year}/{month:02d}/{filename(row)}", data)
                    event(cursor, row['id'], 'monthly_export', actor['username'])
        prefix = 'SPECIMEN_' if mode == 'sandbox-only' else ''
        return file_response(output.getvalue(), f'{prefix}fiches_signees_{year}-{month:02d}.zip', 'application/zip')

    return service
