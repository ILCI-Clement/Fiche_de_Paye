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
from pydantic import BaseModel, ConfigDict, Field

from document_downloads import MAX_ARCHIVE_BYTES, safe_archive_component
from .file_batch import FileBatchAction, DELETE_WORD, cleanup_files
from .signed_fiche_store import SignedFicheStore, ensure_signed_fiche_schema, event, signed_fiche_scope, can_view_signed_fiche
from .role_policy import checked_resource_actor, is_admin, is_super_admin
from .destructive_audit import append_deletion_event

MAX_SIGNED_PDF_BYTES = 25 * 1024 * 1024
PRIVATE_HEADERS = {'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff'}


class SignedFicheDeletion(BaseModel):
    model_config = ConfigDict(extra='forbid')
    confirmation: str = Field(strict=True, min_length=1, max_length=32)


class SignedFileBatch(FileBatchAction):
    ids: list[str] = Field(min_length=1, max_length=100)


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
        if not can_view_signed_fiche(actor, row, target):
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
    def listing(response: Response, year: int = Query(ge=2000, le=2100), month: int | None = Query(None, ge=1, le=12),
                mode: str = Query('production', pattern='^(production|sandbox-only)$'),
                username: str | None = Query(None, max_length=255),
                page: int = Query(1, ge=1), archived: bool = False, actor=Depends(get_current_user)):
        response.headers.update(PRIVATE_HEADERS)
        with service.store.transaction() as cursor:
            actor = current_actor(cursor, actor)
            where = 'fiche_year=%s AND mode=%s AND is_archived=%s'
            args = [year, mode, int(archived)]
            if month is not None:
                where += ' AND fiche_month=%s'
                args.append(month)
            if username is not None:
                target = load_user(cursor, username)
                if not (is_admin(actor) or actor['username'] == username or
                        ('Responsable' in actor['role_tags'] and target and target.get('manager_username') == actor['username'])):
                    raise HTTPException(403, 'Accès aux signatures de cette personne refusé.')
                where += ' AND employee_username=%s'
                args.append(username)
            scope, scope_args = signed_fiche_scope(actor)
            where += f' AND {scope}'
            args += scope_args
            cursor.execute(f'SELECT COUNT(*) AS total FROM signed_fiche_documents WHERE {where}', args)
            total = cursor.fetchone()['total']
            page = min(page, max(1, (total + 19) // 20))
            cursor.execute(f'''SELECT id,document_id,employee_name,fiche_month,fiche_year,mode,status,
                created_at,completed_at,archived_at,file_bytes,last_error,is_archived FROM signed_fiche_documents
                WHERE {where} ORDER BY fiche_month DESC,created_at DESC,id LIMIT 20 OFFSET %s''', [*args, (page - 1) * 20])
            items = cursor.fetchall()
            for item in items:
                item['is_archived'] = bool(item['is_archived'])
        return {'items': items, 'total': total, 'page': page, 'can_export': is_admin(actor), 'can_delete': is_super_admin(actor),
                'automatic_archive': bool(callback_url())}

    def delete_rows(cursor, current, rows):
        paths = []
        for row in rows:
            if row['status'] != 'ready':
                raise HTTPException(409, 'Seuls les PDF signés et archivés peuvent être supprimés.')
            try:
                paths.append(service.path(row))
            except ValueError:
                raise HTTPException(409, 'Le fichier archivé doit être vérifié avant suppression.') from None
        for row in rows:
            append_deletion_event(cursor, current, 'signed_fiche', row['id'],
                                  {'deleted_records': 1, 'document_count': 1, 'confirmed': True, 'previous_is_archived': True})
            event(cursor, row['id'], 'permanently_deleted', current['username'])
            cursor.execute('DELETE FROM signed_fiche_documents WHERE id=%s', (row['id'],))
        return paths

    @app.post('/signed-fiches/batch')
    def batch_signed_files(payload: SignedFileBatch, actor=Depends(get_current_user)):
        ids = sorted(set(payload.ids))
        if any(not re.fullmatch(r'[a-f0-9]{32}', identifier) for identifier in ids):
            raise HTTPException(400, 'Référence de fichier invalide.')
        paths = []
        with service.store.transaction() as cursor:
            current = checked_resource_actor(cursor, actor, load_user, lock=True)
            if not is_admin(current) or (payload.action == 'delete' and not is_super_admin(current)):
                raise HTTPException(403, 'Archivage réservé aux Admin ; suppression réservée au Super Admin.')
            if payload.action == 'delete' and payload.confirmation != DELETE_WORD:
                raise HTTPException(400, 'Saisissez SUPPRIMER pour confirmer.')
            rows = []
            for identifier in ids:
                cursor.execute('SELECT * FROM signed_fiche_documents WHERE id=%s FOR UPDATE', (identifier,))
                row = cursor.fetchone()
                if not row:
                    raise HTTPException(404, 'Un fichier sélectionné est introuvable. Aucune modification effectuée.')
                if row['status'] != 'ready' or bool(row.get('is_archived')) != payload.expected_archived:
                    raise HTTPException(409, 'La sélection a changé ou contient un PDF non archivé. Actualisez la liste.')
                rows.append(row)
            if payload.action == 'delete':
                paths = delete_rows(cursor, current, rows)
            else:
                archived = payload.action == 'archive'
                for row in rows:
                    if bool(row.get('is_archived')) != archived:
                        cursor.execute('UPDATE signed_fiche_documents SET is_archived=%s WHERE id=%s', (int(archived), row['id']))
                        event(cursor, row['id'], 'user_archived' if archived else 'user_restored', current['username'])
        return {'status': 'success', 'count': len(ids), 'cleanup_pending': cleanup_files(paths)}

    @app.delete('/signed-fiches/{archive_id}')
    def delete_signed_fiche(archive_id: str, payload: SignedFicheDeletion, actor=Depends(get_current_user)):
        with service.store.transaction() as cursor:
            current = checked_resource_actor(cursor, actor, load_user, lock=True)
            if not is_super_admin(current):
                raise HTTPException(403, 'La suppression définitive est réservée au Super Admin.')
            cursor.execute('SELECT * FROM signed_fiche_documents WHERE id=%s FOR UPDATE', (archive_id,))
            row = cursor.fetchone()
            if not row:
                raise HTTPException(404, 'Document signé introuvable.')
            if row['status'] != 'ready':
                raise HTTPException(409, 'Seul un PDF signé et archivé peut être supprimé.')
            if payload.confirmation not in (archive_id, DELETE_WORD):
                raise HTTPException(400, 'Saisissez SUPPRIMER pour confirmer.')
            path = delete_rows(cursor, current, [row])[0]
        # Database/audit failure leaves the file intact. A filesystem failure keeps
        # the deleted document inaccessible and is reported explicitly for cleanup.
        try:
            path.unlink(missing_ok=True)
        except OSError:
            logging.getLogger(__name__).warning('Deleted signed fiche file cleanup deferred.')
            return {'status': 'cleanup_pending', 'cleanup_pending': True}
        return {'status': 'success', 'cleanup_pending': False}

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
    def export(year: int = Query(ge=2000, le=2100), month: int | None = Query(None, ge=1, le=12),
               mode: str = Query('production', pattern='^(production|sandbox-only)$'), actor=Depends(get_current_user)):
        with service.store.transaction() as cursor:
            actor = current_actor(cursor, actor)
            if not is_admin(actor):
                raise HTTPException(403, 'Export réservé aux administrateurs.')
            where = "fiche_year=%s AND mode=%s AND status IN ('completed','ready')"
            args = [year, mode]
            if month is not None:
                where += ' AND fiche_month=%s'
                args.append(month)
            cursor.execute(f'''SELECT * FROM signed_fiche_documents WHERE {where}
                ORDER BY fiche_month,employee_name,id LIMIT 5001''', args)
            rows = cursor.fetchall()
            if not rows:
                raise HTTPException(404, 'Aucun PDF signé archivé pour cette période.')
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
                    prefix = 'SPECIMEN_' if mode == 'sandbox-only' else ''
                    folder = f"{year}/{month:02d}" if month is not None else f"{prefix}ILCI_RH_{year}/{int(row['fiche_month']):02d}"
                    archive.writestr(f"{folder}/{filename(row)}", data)
                    event(cursor, row['id'], 'monthly_export' if month is not None else 'yearly_export', actor['username'])
        prefix = 'SPECIMEN_' if mode == 'sandbox-only' else ''
        name = f'{prefix}fiches_signees_{year}-{month:02d}.zip' if month is not None else f'{prefix}ILCI_RH_{year}.zip'
        return file_response(output.getvalue(), name, 'application/zip')

    return service
