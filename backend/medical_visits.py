"""Employee-bound occupational visit dates and private evidence, separate from expenses."""
import base64
import binascii
import json
import os
import secrets
from datetime import date
from io import BytesIO
from pathlib import Path
from urllib.parse import quote
from zipfile import ZIP_DEFLATED, ZipFile
from typing import Annotated, Literal

from document_downloads import MAX_ARCHIVE_BYTES, safe_archive_component

from fastapi import APIRouter, Depends, HTTPException, Query, Response
from pydantic import BaseModel, ConfigDict, Field

from .medical_crypto import MAX_PDF_BYTES, protect, reveal
from .medical_access import may_read_employee, may_edit_employee, readable_usernames, record_archive_filter
from .person_visibility import hidden_account_usernames
from .personnel_register import PersonnelRoute
from .personnel_rules import paris_today
from .role_policy import checked_resource_actor, is_admin


def medical_storage_dir():
    return Path(os.getenv('PRESENCE_MEDICAL_DIR', '/var/lib/presence-app/medical-visits'))


def ensure_medical_schema(connection):
    db = connection()
    try:
        with db.cursor() as c:
            c.execute('''CREATE TABLE IF NOT EXISTS medical_visits (
                id BIGINT NOT NULL AUTO_INCREMENT PRIMARY KEY,
                employee_id BIGINT NOT NULL, visit_date DATE NOT NULL,
                original_filename VARCHAR(255) NULL, stored_filename VARCHAR(100) NULL,
                size_bytes INT NOT NULL DEFAULT 0, version INT NOT NULL DEFAULT 1,
                created_by VARCHAR(255) NOT NULL,
                created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
                KEY medical_employee (employee_id,visit_date,id),
                CONSTRAINT medical_employee_ref FOREIGN KEY (employee_id)
                    REFERENCES personnel_register(employee_id) ON DELETE RESTRICT
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4''')
            c.execute('''CREATE TABLE IF NOT EXISTS medical_visit_events (
                id BIGINT NOT NULL AUTO_INCREMENT PRIMARY KEY,
                visit_id BIGINT NOT NULL, actor_username VARCHAR(255) NOT NULL,
                action VARCHAR(30) NOT NULL, metadata TEXT NOT NULL,
                created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
                KEY medical_event_visit (visit_id,id)
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4''')
            c.execute("SHOW COLUMNS FROM medical_visits LIKE 'is_archived'")
            if not c.fetchone():
                c.execute('ALTER TABLE medical_visits ADD COLUMN is_archived TINYINT(1) NOT NULL DEFAULT 0')
            c.execute('''CREATE TABLE IF NOT EXISTS medical_visit_versions (
                id BIGINT NOT NULL AUTO_INCREMENT PRIMARY KEY,
                visit_id BIGINT NOT NULL, version INT NOT NULL, visit_date DATE NOT NULL,
                original_filename VARCHAR(255) NULL, stored_filename VARCHAR(100) NULL,
                size_bytes INT NOT NULL DEFAULT 0, changed_by VARCHAR(255) NOT NULL,
                created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
                UNIQUE KEY medical_version (visit_id,version),
                CONSTRAINT medical_version_ref FOREIGN KEY (visit_id) REFERENCES medical_visits(id)
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4''')
        db.commit()
    finally:
        db.close()


class PdfInput(BaseModel):
    model_config = ConfigDict(extra='forbid', hide_input_in_errors=True)
    original_filename: str = Field(min_length=1, max_length=255)
    file_b64: str = Field(min_length=1, max_length=14_000_000)


class VisitInput(BaseModel):
    model_config = ConfigDict(extra='forbid', hide_input_in_errors=True)
    employee_id: int = Field(strict=True, ge=1, le=2**63-1)
    visit_date: date
    pdf: PdfInput | None = None


class MedicalAttachment(PdfInput):
    expected_version: int = Field(strict=True, ge=1, le=2**31-2)


class MedicalExportInput(BaseModel):
    model_config = ConfigDict(extra='forbid', hide_input_in_errors=True)
    year: int = Field(strict=True, ge=1900, le=2100)
    month: int | None = Field(default=None, strict=True, ge=1, le=12)
    q: str = Field(default='', max_length=150)
    archived: bool = False
    employee_id: int | None = Field(default=None, strict=True, ge=1, le=2**63-1)
    all_history: bool = False
    format: Literal['zip'] = 'zip'
    visit_ids: list[Annotated[int, Field(strict=True, ge=1, le=2**63-1)]] | None = Field(default=None, min_length=1, max_length=100)


def period_filter(year, month, column='v.visit_date'):
    if month is not None and year is None:
        raise HTTPException(400, 'Choisissez une année avant le mois.')
    if year is None:
        return '', []
    start = date(year, month or 1, 1)
    end = date(year+1, 1, 1) if month in (None, 12) else date(year, month+1, 1)
    return f' AND {column}>=%s AND {column}<%s', [start.isoformat(), end.isoformat()]


def pdf_content(payload):
    name = payload.original_filename.replace('\\', '/').rsplit('/', 1)[-1].strip()
    if not name.lower().endswith('.pdf') or any(ord(ch) < 32 for ch in name):
        raise HTTPException(400, 'Choisissez un fichier PDF.')
    try:
        content = base64.b64decode(payload.file_b64, validate=True)
    except (ValueError, binascii.Error):
        raise HTTPException(400, 'PDF invalide.') from None
    if not content or len(content) > MAX_PDF_BYTES:
        raise HTTPException(400, 'PDF : 10 Mo maximum.')
    # Parse the actual PDF rather than trusting its extension or HTTP type.
    try:
        from io import BytesIO
        from pypdf import PdfReader
        reader = PdfReader(BytesIO(content), strict=True)
        if not content.startswith(b'%PDF-') or reader.is_encrypted or not len(reader.pages):
            raise ValueError()
    except Exception:
        raise HTTPException(400, 'PDF illisible ou protégé par mot de passe.') from None
    return name, content


def register_medical_routes(app, *, get_db_connection, get_current_user, load_user, can_manage=None):
    router = APIRouter(route_class=PersonnelRoute)
    app.router.add_event_handler('startup', lambda: ensure_medical_schema(get_db_connection))

    def employee(c, actor, employee_id, *, write=False):
        c.execute('SELECT employee_id,full_name,linked_username,is_archived FROM personnel_register WHERE employee_id=%s' + (' FOR UPDATE' if write else ''), (employee_id,))
        row = c.fetchone()
        if (not row or not may_read_employee(c, actor, row, load_user, can_manage)
                or row['linked_username'] in hidden_account_usernames(c, actor)):
            raise HTTPException(404, 'Fiche introuvable ou inaccessible.')
        if write and row['is_archived']:
            raise HTTPException(409, 'Cette fiche est archivée.')
        if write and not may_edit_employee(actor, row):
            raise HTTPException(403, 'Vous pouvez modifier uniquement vos propres visites.')
        row['can_edit'] = may_edit_employee(actor, row)
        return row

    def audit(c, actor, visit_id, action, **metadata):
        c.execute('INSERT INTO medical_visit_events (visit_id,actor_username,action,metadata) VALUES(%s,%s,%s,%s)',
                  (visit_id, actor['username'], action, json.dumps(metadata)))

    def scope(c, actor, q, archived, *, mine=False):
        clauses, params = ([], []) if archived else (['e.is_archived=0'], [])
        if not archived:
            clauses.append('''(e.linked_username IS NULL OR EXISTS (
                SELECT 1 FROM users u WHERE u.username=e.linked_username
                AND u.is_active=1 AND u.login_deleted_at IS NULL))''')
        if not is_admin(actor):
            names = [actor['username']] if mine else readable_usernames(c, actor, load_user, can_manage)
            clauses.append('e.linked_username IN (' + ','.join(['%s'] * len(names)) + ')'); params.extend(names)
            if archived:
                clauses.append('(e.is_archived=0 OR e.linked_username=%s)'); params.append(actor['username'])
        hidden = hidden_account_usernames(c, actor)
        if hidden:
            clauses.append('(e.linked_username IS NULL OR e.linked_username NOT IN (' + ','.join(['%s'] * len(hidden)) + '))'); params.extend(hidden)
        if q.strip():
            clauses.append('e.full_name LIKE %s'); params.append('%' + q.strip() + '%')
        return ' WHERE ' + ' AND '.join(clauses or ['1=1']), params

    def read_pdf(visit):
        name = visit['stored_filename']
        if not name: raise HTTPException(404, 'Attestation non déposée.')
        if Path(name).name != name or len(name) != 68:
            raise HTTPException(503, 'Attestation indisponible.')
        try:
            return reveal((medical_storage_dir() / name).read_bytes(), f"{visit['employee_id']}:{name}")
        except Exception:
            raise HTTPException(503, 'Attestation indisponible.') from None

    def store_pdf(c, actor, visit, payload, *, replace=False):
        if visit.get('stored_filename') and not replace:
            raise HTTPException(409, 'Une attestation existe déjà pour cette visite.')
        name, content = pdf_content(payload)
        filename = secrets.token_hex(32) + '.enc'
        try:
            encrypted = protect(content, f"{visit['employee_id']}:{filename}")
        except Exception:
            raise HTTPException(503, 'Le stockage sécurisé des attestations est indisponible.') from None
        folder = medical_storage_dir()
        folder.mkdir(parents=True, exist_ok=True, mode=0o700)
        path = folder / filename
        try:
            with path.open('xb') as stream:
                os.chmod(path, 0o600)
                stream.write(encrypted)
            c.execute('UPDATE medical_visits SET original_filename=%s,stored_filename=%s,size_bytes=%s,version=version+1 WHERE id=%s',
                      (name, filename, len(content), visit['id']))
            audit(c, actor, visit['id'], 'upload')
            return path
        except Exception:
            path.unlink(missing_ok=True)
            raise

    @router.get('/medical-visits/employees')
    def people(q: str = Query('', max_length=150), page: int = Query(1, ge=1, le=1_000_000),
               year: int | None = Query(None, ge=1900, le=2100), month: int | None = Query(None, ge=1, le=12),
               archived: bool = False, mine: bool = False, actor=Depends(get_current_user)):
        period, dates = period_filter(year, month)
        db = get_db_connection()
        try:
            with db.cursor() as c:
                current = checked_resource_actor(c, actor, load_user)
                where, params = scope(c, current, q, archived, mine=mine)
                archive_filter = record_archive_filter(archived)
                c.execute('''SELECT DISTINCT SUBSTR(v.visit_date,1,4) AS year FROM personnel_register e
                    JOIN medical_visits v ON v.employee_id=e.employee_id''' + where + archive_filter + ' ORDER BY year DESC', params)
                years = [int(row['year']) for row in c.fetchall()]
                if period:
                    where += ' AND EXISTS (SELECT 1 FROM medical_visits v WHERE v.employee_id=e.employee_id' + period + archive_filter + ')'
                    params += dates
                elif archived:
                    where += ' AND (e.is_archived=1 OR EXISTS (SELECT 1 FROM medical_visits v WHERE v.employee_id=e.employee_id AND v.is_archived=1))'
                c.execute('SELECT COUNT(*) AS total FROM personnel_register e' + where, params)
                total = c.fetchone()['total']
                c.execute('''SELECT COUNT(*) AS total FROM personnel_register e JOIN medical_visits v ON v.employee_id=e.employee_id'''
                          + where + ' AND v.stored_filename IS NOT NULL' + period + archive_filter, (*params, *dates))
                export_count = c.fetchone()['total']
                c.execute('''SELECT e.employee_id,e.full_name,e.is_archived,v.id AS latest_id,v.visit_date,
                    v.original_filename,(SELECT COUNT(*) FROM medical_visits h WHERE h.employee_id=e.employee_id'''
                    + period.replace('v.visit_date', 'h.visit_date') + record_archive_filter(archived, 'h') + ''') AS visit_count
                    FROM personnel_register e LEFT JOIN medical_visits v ON v.id=(
                        SELECT x.id FROM medical_visits x WHERE x.employee_id=e.employee_id'''
                    + period.replace('v.visit_date', 'x.visit_date') + record_archive_filter(archived, 'x') + ' ORDER BY x.visit_date DESC,x.id DESC LIMIT 1)'
                    + where + ' ORDER BY e.full_name,e.employee_id LIMIT 20 OFFSET %s', (*dates, *dates, *params, (page-1)*20))
                rows = c.fetchall()
                # Only department labels are joined; no contact details or other HR fields leave this API.
                for row in rows:
                    c.execute('''SELECT g.name FROM organization_groups g JOIN user_group_memberships m ON m.group_id=g.id
                        JOIN personnel_register e ON e.linked_username=m.username
                        WHERE e.employee_id=%s AND m.relation='member' AND g.is_active=1 ORDER BY g.name''', (row['employee_id'],))
                    row['departments'] = [g['name'] for g in c.fetchall()]
                return {'employees': rows, 'total': total, 'years': years, 'export_count': export_count}
        finally:
            db.close()

    @router.get('/medical-visits/employees/{employee_id}')
    def history(employee_id: int, actor=Depends(get_current_user), page: int = Query(1, ge=1, le=1_000_000),
                year: int | None = Query(None, ge=1900, le=2100), month: int | None = Query(None, ge=1, le=12), archived: bool = False):
        period, dates = period_filter(year, month, 'visit_date')
        db = get_db_connection()
        try:
            with db.cursor() as c:
                current = checked_resource_actor(c, actor, load_user)
                person = employee(c, current, employee_id)
                record_filter = ' AND is_archived=%s'
                # Personnel archive view includes all retained visits for an archived dossier.
                archive_params = [int(archived)]
                if archived and person['is_archived']:
                    record_filter, archive_params = '', []
                c.execute('SELECT COUNT(*) AS total FROM medical_visits WHERE employee_id=%s' + period + record_filter, (employee_id, *dates, *archive_params))
                total = c.fetchone()['total']
                c.execute('''SELECT id,visit_date,original_filename,size_bytes,version,created_at,is_archived
                    FROM medical_visits WHERE employee_id=%s''' + period + record_filter + ' ORDER BY visit_date DESC,id DESC LIMIT 20 OFFSET %s', (employee_id, *dates, *archive_params, (page-1)*20))
                return {'employee': person, 'visits': c.fetchall(), 'total': total}
        finally:
            db.close()

    @router.post('/medical-visits')
    def create(payload: VisitInput, actor=Depends(get_current_user)):
        if payload.visit_date > paris_today():
            raise HTTPException(400, 'La date de visite ne peut pas être future.')
        db, path = get_db_connection(), None
        try:
            with db.cursor() as c:
                current = checked_resource_actor(c, actor, load_user, lock=True)
                employee(c, current, payload.employee_id, write=True)
                c.execute('INSERT INTO medical_visits (employee_id,visit_date,created_by) VALUES(%s,%s,%s)',
                          (payload.employee_id, payload.visit_date, current['username']))
                identity = c.lastrowid
                audit(c, current, identity, 'create', visit_date=payload.visit_date.isoformat())
                if payload.pdf:
                    path = store_pdf(c, current, {'id': identity, 'employee_id': payload.employee_id}, payload.pdf)
            db.commit()
            return {'id': identity}
        except Exception:
            db.rollback()
            if path: path.unlink(missing_ok=True)
            raise
        finally:
            db.close()

    @router.post('/medical-visits/{visit_id}/file')
    def attach(visit_id: int, payload: MedicalAttachment, actor=Depends(get_current_user)):
        db, path = get_db_connection(), None
        try:
            with db.cursor() as c:
                current = checked_resource_actor(c, actor, load_user, lock=True)
                c.execute('SELECT * FROM medical_visits WHERE id=%s FOR UPDATE', (visit_id,))
                visit = c.fetchone()
                if not visit: raise HTTPException(404, 'Visite introuvable.')
                employee(c, current, visit['employee_id'], write=True)
                if visit['version'] != payload.expected_version:
                    raise HTTPException(409, 'Cette visite a changé. Actualisez avant de réessayer.')
                if visit['is_archived']: raise HTTPException(409, 'Restaurez cette visite avant modification.')
                path = store_pdf(c, current, visit, payload)
            db.commit()
            return {'status': 'success'}
        except Exception:
            db.rollback()
            if path: path.unlink(missing_ok=True)
            raise
        finally:
            db.close()

    @router.get('/medical-visits/{visit_id}/file')
    def file(visit_id: int, actor=Depends(get_current_user)):
        db = get_db_connection()
        try:
            with db.cursor() as c:
                current = checked_resource_actor(c, actor, load_user, lock=True)
                c.execute('SELECT * FROM medical_visits WHERE id=%s', (visit_id,))
                visit = c.fetchone()
                if not visit: raise HTTPException(404, 'Visite introuvable.')
                employee(c, current, visit['employee_id'])
                content = read_pdf(visit)
                audit(c, current, visit_id, 'read')
            db.commit()
            return Response(content, media_type='application/pdf', headers={
                'Content-Disposition': "attachment; filename*=UTF-8''" + quote(visit['original_filename'], safe=''),
                'X-Content-Type-Options': 'nosniff'})
        except Exception:
            db.rollback()
            raise
        finally:
            db.close()

    @router.get('/medical-visits/{visit_id}')
    def linked_visit(visit_id: int, actor=Depends(get_current_user)):
        db = get_db_connection()
        try:
            with db.cursor() as c:
                current = checked_resource_actor(c, actor, load_user, lock=True)
                c.execute('SELECT * FROM medical_visits WHERE id=%s', (visit_id,))
                visit = c.fetchone()
                if not visit: raise HTTPException(404, 'Visite introuvable.')
                employee(c, current, visit['employee_id'])
                return {key: visit[key] for key in ('id', 'visit_date', 'original_filename', 'size_bytes', 'version', 'created_at', 'is_archived')}
        finally:
            db.rollback()
            db.close()

    @router.post('/medical-visits/export')
    def export(payload: MedicalExportInput, actor=Depends(get_current_user)):
        db = get_db_connection()
        try:
            with db.cursor() as c:
                current = checked_resource_actor(c, actor, load_user, lock=True)
                where, params = scope(c, current, payload.q, payload.archived)
                if payload.employee_id is not None:
                    employee(c, current, payload.employee_id)
                    where += ' AND e.employee_id=%s'; params.append(payload.employee_id)
                if payload.all_history and payload.employee_id is None:
                    raise HTTPException(400, 'Choisissez une personne pour exporter son historique complet.')
                period, dates = ('', []) if payload.all_history else period_filter(payload.year, payload.month)
                selection, selected_params = '', []
                if payload.visit_ids is not None:
                    selected_params = list(set(payload.visit_ids))
                    selection = ' AND v.id IN (' + ','.join(['%s'] * len(selected_params)) + ')'
                c.execute('''SELECT v.*,e.full_name FROM medical_visits v
                    JOIN personnel_register e ON e.employee_id=v.employee_id'''
                    + where + ' AND v.stored_filename IS NOT NULL' + period + ('' if payload.all_history else record_archive_filter(payload.archived)) + selection + ' ORDER BY v.visit_date,v.id LIMIT 101', (*params, *dates, *selected_params))
                visits = c.fetchall()
                if selected_params and {v['id'] for v in visits} != set(selected_params):
                    raise HTTPException(404, 'Sélection indisponible. Actualisez la liste.')
                if not visits:
                    raise HTTPException(400, 'Aucune attestation PDF pour cette période.')
                if len(visits) > 100 or sum(v['size_bytes'] for v in visits) > MAX_ARCHIVE_BYTES:
                    raise HTTPException(400, 'Export limité à 100 PDF / 100 Mo. Filtrez par mois ou par personne.')
                output, total = BytesIO(), 0
                with ZipFile(output, 'w', compression=ZIP_DEFLATED) as archive:
                    for visit in visits:
                        content = read_pdf(visit)
                        total += len(content)
                        if total > MAX_ARCHIVE_BYTES:
                            raise HTTPException(400, 'Export limité à 100 Mo. Aucun ZIP partiel n’a été créé.')
                        day = str(visit['visit_date'])[:10]
                        person = safe_archive_component(visit['full_name']) + f"_{visit['employee_id']}"
                        filename = f"{person}_{day}_{visit['id']}.pdf"
                        archive.writestr(f'{day[:4]}/{day[5:7]}/{filename}', content)
                        audit(c, current, visit['id'], 'export', format=payload.format,
                              year=None if payload.all_history else payload.year,
                              month=None if payload.all_history else payload.month, all_history=payload.all_history)
                if output.getbuffer().nbytes > MAX_ARCHIVE_BYTES:
                    raise HTTPException(400, 'Export limité à 100 Mo.')
            db.commit()
            suffix = (f"salarie-{payload.employee_id}_historique" if payload.all_history else
                      str(payload.year) + (f'-{payload.month:02d}' if payload.month else ''))
            return Response(output.getvalue(), media_type='application/zip', headers={
                'Content-Disposition': f'attachment; filename="visites-medicales_{suffix}.zip"',
                'X-Content-Type-Options': 'nosniff'})
        except Exception:
            db.rollback()
            raise
        finally:
            db.close()

    from .medical_history import register_medical_history
    register_medical_history(router, get_db_connection=get_db_connection, get_current_user=get_current_user,
                             load_user=load_user, employee=employee, audit=audit, read_pdf=read_pdf,
                             store_pdf=store_pdf)
    app.include_router(router)
