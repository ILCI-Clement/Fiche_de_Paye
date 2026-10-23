"""Permanent signature metadata. No callback credentials or PDFs in public DTOs."""
import hashlib
import re
import secrets
import uuid
from contextlib import contextmanager

from fastapi import HTTPException


def ensure_signed_fiche_schema(cursor):
    cursor.execute('''CREATE TABLE IF NOT EXISTS signed_fiche_documents (
        id CHAR(32) NOT NULL PRIMARY KEY,
        reference_hash CHAR(64) NOT NULL UNIQUE,
        namespace VARCHAR(128) NOT NULL,
        document_id VARCHAR(128) NULL UNIQUE,
        employee_username VARCHAR(255) NOT NULL,
        sender_username VARCHAR(255) NOT NULL,
        employee_name VARCHAR(255) NOT NULL,
        fiche_month INT NOT NULL, fiche_year INT NOT NULL,
        mode VARCHAR(32) NOT NULL,
        status VARCHAR(32) NOT NULL DEFAULT 'pending',
        signed_pdf_url TEXT NULL,
        completed_at DATETIME(6) NULL, archived_at DATETIME(6) NULL,
        sha256 CHAR(64) NULL, file_bytes BIGINT NULL,
        download_attempts INT NOT NULL DEFAULT 0,
        next_attempt_at DATETIME(6) NULL,
        lease_until DATETIME(6) NULL, lease_token CHAR(32) NULL,
        last_error VARCHAR(255) NULL,
        created_at DATETIME(6) NOT NULL DEFAULT CURRENT_TIMESTAMP(6),
        KEY signed_fiche_period (fiche_year, fiche_month, mode),
        KEY signed_fiche_employee (employee_username),
        KEY signed_fiche_pending (status, next_attempt_at)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS signed_fiche_events (
        id BIGINT NOT NULL AUTO_INCREMENT PRIMARY KEY,
        archive_id CHAR(32) NOT NULL,
        event VARCHAR(64) NOT NULL,
        actor_username VARCHAR(255) NULL,
        created_at DATETIME(6) NOT NULL DEFAULT CURRENT_TIMESTAMP(6),
        KEY signed_fiche_event_document (archive_id, id)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4''')


def rename_signed_fiche_owner(cursor, old, new):
    # Same transaction as the account rename; original signer names stay historical.
    for column in ('employee_username', 'sender_username'):
        cursor.execute(f'UPDATE signed_fiche_documents SET {column}=%s WHERE {column}=%s', (new, old))


def event(cursor, archive_id, name, actor=None):
    cursor.execute('INSERT INTO signed_fiche_events (archive_id,event,actor_username) VALUES (%s,%s,%s)',
                   (archive_id, name, actor))


class SignedFicheStore:
    def __init__(self, connect):
        self.connect = connect

    @contextmanager
    def transaction(self):
        db = self.connect()
        try:
            with db.cursor() as cursor:
                yield cursor
            db.commit()
        except Exception:
            db.rollback()
            raise
        finally:
            db.close()

    def prepare(self, *, employee, sender, name, month, year, namespace, mode):
        archive_id = uuid.uuid4().hex
        # A per-request 256-bit callback capability, carried in the POST body only.
        # Keep only its hash; never put the capability in URLs, UI or access logs.
        reference = f'presence-{year}-{month:02d}-{secrets.token_urlsafe(32)}'
        with self.transaction() as cursor:
            cursor.execute('''INSERT INTO signed_fiche_documents
                (id,reference_hash,namespace,employee_username,sender_username,employee_name,
                 fiche_month,fiche_year,mode) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)''',
                (archive_id, hashlib.sha256(reference.encode()).hexdigest(), namespace,
                 employee, sender, name, month, year, mode))
            event(cursor, archive_id, 'prepared', sender)
        return archive_id, reference

    def accepted(self, archive_id, document_id):
        with self.transaction() as cursor:
            cursor.execute('SELECT document_id,status FROM signed_fiche_documents WHERE id=%s FOR UPDATE', (archive_id,))
            row = cursor.fetchone()
            if row['document_id'] and row['document_id'] != document_id:
                raise HTTPException(409, 'La référence de signature doit être vérifiée.')
            cursor.execute('''UPDATE signed_fiche_documents SET document_id=%s,
                status=IF(status IN ('pending','uncertain'),'sent',status) WHERE id=%s''', (document_id, archive_id))
            event(cursor, archive_id, 'accepted')

    def uncertain(self, archive_id):
        with self.transaction() as cursor:
            cursor.execute("UPDATE signed_fiche_documents SET status='uncertain' WHERE id=%s AND status='pending'", (archive_id,))

    def receive(self, payload, validate_url):
        reference = payload.get('reference_id')
        if not isinstance(reference, str) or len(reference) > 256:
            raise HTTPException(404, 'Notification introuvable.')
        with self.transaction() as cursor:
            cursor.execute('SELECT * FROM signed_fiche_documents WHERE reference_hash=%s FOR UPDATE',
                           (hashlib.sha256(reference.encode()).hexdigest(),))
            row = cursor.fetchone()
            if not row or payload.get('namespace') != row['namespace'] or payload.get('provider') != 'clawshow_esign':
                raise HTTPException(404, 'Notification introuvable.')
            document_id = payload.get('document_id')
            if not isinstance(document_id, str) or not re.fullmatch(r'[A-Za-z0-9_-]{1,128}', document_id):
                raise HTTPException(400, 'Référence de document invalide.')
            if row['document_id'] and document_id != row['document_id']:
                raise HTTPException(409, 'Référence de document différente.')
            kind = payload.get('event')
            if not isinstance(kind, str):
                raise HTTPException(400, 'Événement invalide.')
            if kind not in {'document.created', 'signer.signed', 'document.completed', 'signer.declined',
                            'document.expired', 'document.cancelled'}:
                return
            # Completion wins over delayed/duplicate intermediate notifications.
            if row['status'] in {'completed', 'ready'}:
                return
            complete = kind in {'document.completed', 'signer.signed'} and payload.get('all_signed') is True
            if kind == 'document.completed' and not complete:
                raise HTTPException(400, 'La confirmation de toutes les signatures est requise.')
            url = validate_url(payload.get('signed_pdf_url'), document_id) if complete else None
            status = ('completed' if complete else
                      {'signer.signed': 'partially_signed', 'signer.declined': 'declined',
                       'document.expired': 'expired', 'document.cancelled': 'cancelled'}.get(kind, row['status']))
            if row['status'] in {'declined', 'expired', 'cancelled'} and not complete:
                return
            cursor.execute('''UPDATE signed_fiche_documents SET document_id=%s,status=%s,
                signed_pdf_url=COALESCE(%s,signed_pdf_url),
                completed_at=IF(%s,UTC_TIMESTAMP(6),completed_at),
                next_attempt_at=IF(%s,UTC_TIMESTAMP(6),next_attempt_at) WHERE id=%s''',
                (document_id, status, url, complete, complete, row['id']))
            if status != row['status']:
                event(cursor, row['id'], kind)

    def claim(self):
        with self.transaction() as cursor:
            cursor.execute('''SELECT * FROM signed_fiche_documents WHERE status='completed'
                AND next_attempt_at<=UTC_TIMESTAMP(6)
                AND (lease_until IS NULL OR lease_until<UTC_TIMESTAMP(6))
                ORDER BY next_attempt_at,id LIMIT 1 FOR UPDATE''')
            row = cursor.fetchone()
            if not row:
                return None
            row['lease_token'] = uuid.uuid4().hex
            cursor.execute('''UPDATE signed_fiche_documents SET lease_token=%s,
                lease_until=DATE_ADD(UTC_TIMESTAMP(6),INTERVAL 180 SECOND),
                download_attempts=download_attempts+1 WHERE id=%s''', (row['lease_token'], row['id']))
            return row

    def downloaded(self, row, digest, size):
        with self.transaction() as cursor:
            cursor.execute('''UPDATE signed_fiche_documents SET status='ready',sha256=%s,file_bytes=%s,
                archived_at=UTC_TIMESTAMP(6),last_error=NULL,signed_pdf_url=NULL,
                lease_until=NULL,lease_token=NULL WHERE id=%s AND lease_token=%s AND status='completed' ''',
                (digest, size, row['id'], row['lease_token']))
            if cursor.rowcount:
                event(cursor, row['id'], 'archived')

    def failed(self, row):
        delay = min(3600, 30 * 2 ** min(row['download_attempts'], 7))
        with self.transaction() as cursor:
            cursor.execute('''UPDATE signed_fiche_documents SET last_error=%s,
                next_attempt_at=DATE_ADD(UTC_TIMESTAMP(6),INTERVAL %s SECOND),
                lease_until=NULL,lease_token=NULL WHERE id=%s AND lease_token=%s AND status='completed' ''',
                ('Téléchargement différé ; nouvelle tentative automatique.', delay, row['id'], row['lease_token']))
            if cursor.rowcount:
                event(cursor, row['id'], 'download_deferred')
