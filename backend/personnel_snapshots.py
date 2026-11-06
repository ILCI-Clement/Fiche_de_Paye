"""Immutable encrypted XLSX snapshots. No plaintext files, deletes or historical reconstruction."""
from contextlib import contextmanager
from datetime import timedelta
import hashlib
import json
import os
import threading
import uuid

from fastapi import Depends, HTTPException, Query
from fastapi.responses import Response
from pydantic import BaseModel, ConfigDict, Field, StrictBool
from typing import Literal

from .personnel_export import generate_personnel_excel, MAX_EXPORT_ROWS
from .personnel_pdf import personnel_excel_to_pdf
from .personnel_sensitive import decrypt_social_security_number
from .personnel_store import PERIOD_FIELDS
from .role_policy import checked_resource_actor, is_admin, is_super_admin
from .snapshot_schedule import next_due, utc_now
from .snapshot_crypto import keys, protect, reveal

PUBLIC_COLUMNS = 'id,source,status,scheduled_for,captured_at,completed_at,created_at,last_attempt,employee_count,row_count,error_code,actor_username'


def ensure_snapshot_schema(cursor, now=None):
    cursor.execute('''CREATE TABLE IF NOT EXISTS personnel_snapshot_settings (
        id INT PRIMARY KEY,interval_months INT NOT NULL,version INT NOT NULL DEFAULT 1,
        next_due DATETIME NOT NULL,updated_by VARCHAR(255) NULL,updated_at DATETIME NOT NULL
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS personnel_snapshots (
        id CHAR(36) PRIMARY KEY,source VARCHAR(16) NOT NULL,status VARCHAR(16) NOT NULL,
        scheduled_for DATETIME NULL UNIQUE,captured_at DATETIME NULL,completed_at DATETIME NULL,
        created_at DATETIME NOT NULL,last_attempt DATETIME NULL,actor_username VARCHAR(255) NULL,
        employee_count INT NOT NULL DEFAULT 0,row_count INT NOT NULL DEFAULT 0,
        error_code VARCHAR(64) NOT NULL DEFAULT '',
        full_encrypted LONGBLOB NULL,redacted_encrypted LONGBLOB NULL,
        full_sha256 CHAR(64) NULL,redacted_sha256 CHAR(64) NULL,
        KEY snapshot_order(created_at,id)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS personnel_snapshot_events (
        id BIGINT AUTO_INCREMENT PRIMARY KEY,snapshot_id CHAR(36) NULL,action VARCHAR(32) NOT NULL,
        actor_username VARCHAR(255) NULL,details TEXT NOT NULL,created_at DATETIME NOT NULL
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4''')
    now = now or utc_now()
    cursor.execute('''INSERT IGNORE INTO personnel_snapshot_settings
        (id,interval_months,next_due,updated_at) VALUES(1,6,%s,%s)''', (next_due(now, 6), now))


def event(cursor, identity, action, actor=None, **details):
    cursor.execute('''INSERT INTO personnel_snapshot_events(snapshot_id,action,actor_username,details,created_at)
        VALUES(%s,%s,%s,%s,%s)''', (identity, action, actor, json.dumps(details), utc_now()))


class ScheduleInput(BaseModel):
    model_config = ConfigDict(extra='forbid')
    interval_months: Literal[1, 3, 6, 12]
    expected_version: int = Field(strict=True, ge=1)


class DownloadInput(BaseModel):
    model_config = ConfigDict(extra='forbid')
    include_nir: StrictBool = False
    confirm_sensitive: StrictBool = False
    format: Literal['pdf', 'xlsx'] = 'pdf'


class ManualInput(BaseModel):
    model_config = ConfigDict(extra='forbid')
    request_id: uuid.UUID


class SnapshotService:
    def __init__(self, connect, load_user):
        self.connect, self.load_user = connect, load_user

    def admin(self, cursor, actor, *, lock=False):
        current = checked_resource_actor(cursor, actor, self.load_user, lock=lock)
        if not is_admin(current):
            raise HTTPException(403, 'Archives réservées aux administrateurs.')
        return current

    @contextmanager
    def writer(self):
        db = self.connect()
        acquired = False
        try:
            with db.cursor() as cursor:
                cursor.execute("SELECT GET_LOCK(CONCAT('rh-snapshots:',DATABASE()),0) AS acquired")
                acquired = bool(cursor.fetchone()['acquired'])
            if not acquired:
                raise HTTPException(409, 'Une sauvegarde du registre est déjà en cours.')
            yield db
        except Exception:
            db.rollback()
            raise
        finally:
            db.rollback()  # Release locks/read snapshots on early idempotent returns too.
            if acquired:
                with db.cursor() as cursor:
                    cursor.execute("SELECT RELEASE_LOCK(CONCAT('rh-snapshots:',DATABASE()))")
            db.close()

    def get(self, cursor, identity):
        cursor.execute(f'SELECT {PUBLIC_COLUMNS} FROM personnel_snapshots WHERE id=%s', (identity,))
        row = cursor.fetchone()
        if not row:
            raise HTTPException(404, 'Archive introuvable.')
        return row

    def capture(self, db, identity):
        """Read all fields/periods in one consistent snapshot, then atomically publish both variants."""
        try:
            keys()  # Fail closed before reading sensitive source rows.
            with db.cursor() as cursor:
                cursor.execute('SET TRANSACTION ISOLATION LEVEL REPEATABLE READ')
                cursor.execute('START TRANSACTION WITH CONSISTENT SNAPSHOT')
                captured = utc_now()
                cursor.execute('''SELECT e.employee_id,e.full_name,e.nationality,e.sex,e.address,e.birth_date,
                    e.job_title,e.qualification,e.social_security_number_encrypted,'''
                    + ','.join(f'h.{column}' for column in PERIOD_FIELDS)
                    + ''' FROM personnel_register e LEFT JOIN personnel_register_periods h
                    ON h.employee_id=e.employee_id WHERE e.is_archived=0
                    ORDER BY e.employee_id,h.entry_date,h.period_id LIMIT %s''', (MAX_EXPORT_ROWS + 1,))
                rows = cursor.fetchall()
            db.commit()
            if len(rows) > MAX_EXPORT_ROWS:
                raise ValueError('row_limit')
            numbers = {}
            for row in rows:
                employee = row['employee_id']
                if employee not in numbers:
                    envelope = row['social_security_number_encrypted']
                    numbers[employee] = decrypt_social_security_number(envelope, employee) if envelope is not None else ''
                row['social_security_number'] = numbers[employee]
                row.pop('social_security_number_encrypted', None)
            full = generate_personnel_excel(rows, allow_empty=True, snapshot_at=captured)
            for row in rows:
                row['social_security_number'] = ''
            redacted = generate_personnel_excel(rows, allow_empty=True, snapshot_at=captured)
            full_encrypted = protect(full, identity, 'full')
            redacted_encrypted = protect(redacted, identity, 'redacted')
            with db.cursor() as cursor:
                cursor.execute('''UPDATE personnel_snapshots SET status='ready',captured_at=%s,completed_at=%s,
                    employee_count=%s,row_count=%s,error_code='',full_encrypted=%s,redacted_encrypted=%s,
                    full_sha256=%s,redacted_sha256=%s WHERE id=%s AND status='running' ''',
                    (captured, utc_now(), len(numbers), len(rows), full_encrypted, redacted_encrypted,
                     hashlib.sha256(full).hexdigest(), hashlib.sha256(redacted).hexdigest(), identity))
                if cursor.rowcount != 1:
                    raise RuntimeError('snapshot_publish_conflict')
                event(cursor, identity, 'completed', employee_count=len(numbers), row_count=len(rows))
            db.commit()
        except Exception as error:
            db.rollback()
            code = 'row_limit' if isinstance(error, ValueError) and str(error) == 'row_limit' else 'generation_failed'
            # No exception text, names, source values, numbers, paths or keys in this audit.
            with db.cursor() as cursor:
                cursor.execute("UPDATE personnel_snapshots SET status='failed',error_code=%s WHERE id=%s AND status='running'", (code, identity))
                event(cursor, identity, 'failed', error_code=code)
            db.commit()
        with db.cursor() as cursor:
            return self.get(cursor, identity)

    def create(self, actor=None, identity=None, retry=None):
        with self.writer() as db:
            now = utc_now()
            with db.cursor() as cursor:
                current = self.admin(cursor, actor, lock=True) if actor else None
                cursor.execute('SELECT * FROM personnel_snapshot_settings WHERE id=1 FOR UPDATE')
                settings = cursor.fetchone()
                if retry:
                    row = self.get(cursor, retry)
                    if row['status'] == 'ready':
                        return row
                    identity = retry
                elif identity:
                    cursor.execute(f'SELECT {PUBLIC_COLUMNS} FROM personnel_snapshots WHERE id=%s', (identity,))
                    prior = cursor.fetchone()
                    if prior:
                        return prior
                    cursor.execute('''INSERT INTO personnel_snapshots(id,source,status,created_at,actor_username)
                        VALUES(%s,'manual','running',%s,%s)''', (identity, now, current['username']))
                else:
                    cursor.execute("SELECT id FROM personnel_snapshots WHERE source<>'manual' AND (status='running' OR (status='failed' AND last_attempt<=%s)) ORDER BY created_at LIMIT 1", (now - timedelta(hours=1),))
                    failed = cursor.fetchone()
                    if failed:
                        identity = failed['id']
                    elif settings['next_due'] <= now:
                        identity = str(uuid.uuid4())
                        source = 'catchup' if now - settings['next_due'] > timedelta(minutes=5) else 'scheduled'
                        cursor.execute('''INSERT INTO personnel_snapshots(id,source,status,scheduled_for,created_at)
                            VALUES(%s,%s,'running',%s,%s)''', (identity, source, settings['next_due'], now))
                    else:
                        return None
                    if settings['next_due'] <= now:
                        cursor.execute('UPDATE personnel_snapshot_settings SET next_due=%s WHERE id=1', (next_due(now, settings['interval_months']),))
                cursor.execute("UPDATE personnel_snapshots SET status='running',last_attempt=%s WHERE id=%s AND status<>'ready'", (now, identity))
                event(cursor, identity, 'retry' if retry else 'started', current['username'] if current else None)
            db.commit()
            return self.capture(db, identity)


def register_snapshot_routes(app, *, get_db_connection, get_current_user, load_user):
    service = SnapshotService(get_db_connection, load_user)
    app.state.personnel_snapshots = service
    stop = threading.Event()

    @app.on_event('startup')
    def start():
        db = get_db_connection()
        try:
            with db.cursor() as cursor:
                ensure_snapshot_schema(cursor)
            db.commit()
        finally:
            db.close()
        if os.getenv('PRESENCE_PERSONNEL_SNAPSHOT_WORKER', '1') == '1':
            def worker():
                while not stop.is_set():
                    try:
                        service.create()
                    except Exception:
                        pass  # Database outage: retry later without logging sensitive exceptions.
                    stop.wait(60)
            threading.Thread(target=worker, daemon=True, name='personnel-snapshots').start()

    @app.on_event('shutdown')
    def shutdown():
        stop.set()

    @app.middleware('http')
    async def private_archives(request, call_next):
        response = await call_next(request)
        if request.url.path.startswith('/personnel-snapshots'):
            response.headers['Cache-Control'] = 'no-store, private'
            response.headers['Pragma'] = 'no-cache'
        return response

    @app.get('/personnel-snapshots')
    def listing(page: int = Query(1, ge=1), actor=Depends(get_current_user)):
        db = get_db_connection()
        try:
            with db.cursor() as cursor:
                current = service.admin(cursor, actor)
                cursor.execute('SELECT * FROM personnel_snapshot_settings WHERE id=1')
                settings = cursor.fetchone()
                cursor.execute('SELECT COUNT(*) AS total FROM personnel_snapshots')
                total = cursor.fetchone()['total']
                cursor.execute(f'SELECT {PUBLIC_COLUMNS} FROM personnel_snapshots ORDER BY created_at DESC,id DESC LIMIT 20 OFFSET %s', ((page - 1) * 20,))
                rows = cursor.fetchall()
                try:
                    keys(); protected = True
                except Exception:
                    protected = False
                return {'settings': settings, 'snapshots': rows, 'total': total,
                        'can_configure': is_super_admin(current), 'can_export_excel': is_super_admin(current), 'encryption_ready': protected}
        finally:
            db.close()

    @app.put('/personnel-snapshots/settings')
    def configure(payload: ScheduleInput, actor=Depends(get_current_user)):
        with service.writer() as db:
            with db.cursor() as cursor:
                current = service.admin(cursor, actor, lock=True)
                if not is_super_admin(current):
                    raise HTTPException(403, 'Seul un Super Admin peut modifier le cycle.')
                cursor.execute('SELECT * FROM personnel_snapshot_settings WHERE id=1 FOR UPDATE')
                previous = cursor.fetchone()
                if previous['version'] != payload.expected_version:
                    raise HTTPException(409, 'Le cycle a changé. Actualisez la fenêtre.')
                now = utc_now()
                # Preserve an already due run so changing the cycle cannot erase a missed archive.
                due = previous['next_due'] if previous['next_due'] <= now else next_due(now, payload.interval_months)
                cursor.execute('''UPDATE personnel_snapshot_settings SET interval_months=%s,version=version+1,
                    next_due=%s,updated_at=%s,updated_by=%s WHERE id=1''', (payload.interval_months, due, now, current['username']))
                event(cursor, None, 'schedule_changed', current['username'], before=previous['interval_months'], after=payload.interval_months)
            db.commit()
        return {'status': 'success'}

    @app.post('/personnel-snapshots')
    def manual(payload: ManualInput, actor=Depends(get_current_user)):
        return service.create(actor, str(payload.request_id))

    @app.post('/personnel-snapshots/{identity}/retry')
    def retry(identity: uuid.UUID, actor=Depends(get_current_user)):
        return service.create(actor, retry=str(identity))

    @app.post('/personnel-snapshots/{identity}/download')
    def download(identity: uuid.UUID, payload: DownloadInput, actor=Depends(get_current_user)):
        if payload.include_nir and not payload.confirm_sensitive:
            raise HTTPException(400, 'Confirmez l’inclusion des numéros de sécurité sociale.')
        db = get_db_connection()
        try:
            with db.cursor() as cursor:
                current = service.admin(cursor, actor, lock=True)
                if payload.format == 'xlsx' and not is_super_admin(current):
                    raise HTTPException(403, 'Le téléchargement Excel est réservé au Super Admin.')
                row = service.get(cursor, str(identity))
                if row['status'] != 'ready':
                    raise HTTPException(409, 'Cette archive n’est pas encore disponible.')
                variant = 'full' if payload.include_nir else 'redacted'
                cursor.execute(f'SELECT {variant}_encrypted AS content,{variant}_sha256 AS digest FROM personnel_snapshots WHERE id=%s', (str(identity),))
                encrypted = cursor.fetchone()
                try:
                    content = reveal(bytes(encrypted['content']), str(identity), variant)
                    if hashlib.sha256(content).hexdigest() != encrypted['digest']:
                        raise ValueError('invalid_digest')
                except Exception:
                    raise HTTPException(503, 'Archive indisponible : vérifiez la protection et les clés historiques.') from None
                if payload.format == 'pdf':
                    try:
                        content = personnel_excel_to_pdf(content, snapshot_at=row['captured_at'])
                    except Exception:
                        raise HTTPException(503, 'La génération du PDF est indisponible. Aucun fichier fourni.') from None
                event(cursor, str(identity), 'downloaded', current['username'], include_nir=payload.include_nir, format=payload.format)
            db.commit()
            suffix = 'avec-nir' if payload.include_nir else 'sans-nir'
            stamp = row['captured_at'].strftime('%Y%m%dT%H%M%SZ')
            return Response(content, media_type='application/pdf' if payload.format == 'pdf' else 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                            headers={'Content-Disposition': f'attachment; filename="utilisateurs-{stamp}-{identity}-{suffix}.{payload.format}"'})
        except Exception:
            db.rollback()
            raise
        finally:
            db.close()

    return service
