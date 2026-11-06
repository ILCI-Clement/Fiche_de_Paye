"""Versioned corrections, administrator archiving and guarded Super Admin deletion."""
from datetime import date
from pathlib import Path
from urllib.parse import quote

from fastapi import Depends, HTTPException, Query, Response
from pydantic import BaseModel, ConfigDict, Field, StrictBool

from .destructive_audit import append_deletion_event
from .file_batch import DELETE_WORD, cleanup_files
from .medical_visits import PdfInput, medical_storage_dir
from .personnel_rules import paris_today
from .role_policy import checked_resource_actor, is_admin, is_super_admin


class VisitVersionInput(BaseModel):
    model_config = ConfigDict(extra='forbid', hide_input_in_errors=True)
    expected_version: int = Field(strict=True, ge=1, le=2**31-2)


class VisitCorrection(VisitVersionInput):
    visit_date: date
    pdf: PdfInput | None = None


class VisitArchive(VisitVersionInput):
    is_archived: StrictBool


class VisitDelete(VisitVersionInput):
    confirmation: str = Field(strict=True, max_length=32)


PUBLIC_FIELDS = ('id', 'visit_date', 'original_filename', 'size_bytes', 'version', 'created_at', 'is_archived')


def register_medical_history(router, *, get_db_connection, get_current_user, load_user,
                             employee, audit, read_pdf, store_pdf):
    def load(c, visit_id, *, lock=False):
        c.execute('SELECT * FROM medical_visits WHERE id=%s' + (' FOR UPDATE' if lock else ''), (visit_id,))
        visit = c.fetchone()
        if not visit: raise HTTPException(404, 'Visite introuvable.')
        return visit

    def version_matches(visit, payload):
        if visit['version'] != payload.expected_version:
            raise HTTPException(409, 'Cette visite a changé. Actualisez avant de réessayer.')

    @router.get('/medical-visits/{visit_id}/versions')
    def versions(visit_id: int, page: int = Query(1, ge=1, le=1_000_000), actor=Depends(get_current_user)):
        db = get_db_connection()
        try:
            with db.cursor() as c:
                current = checked_resource_actor(c, actor, load_user)
                visit = load(c, visit_id)
                employee(c, current, visit['employee_id'])
                c.execute('SELECT COUNT(*) AS total FROM medical_visit_versions WHERE visit_id=%s', (visit_id,))
                total = c.fetchone()['total']
                c.execute('''SELECT version,visit_date,original_filename,size_bytes,created_at
                    FROM medical_visit_versions WHERE visit_id=%s ORDER BY version DESC LIMIT 20 OFFSET %s''', (visit_id, (page-1)*20))
                return {'current': {key: visit[key] for key in PUBLIC_FIELDS}, 'versions': c.fetchall(), 'total': total}
        finally:
            db.close()

    @router.get('/medical-visits/{visit_id}/versions/{version}/file')
    def version_file(visit_id: int, version: int, actor=Depends(get_current_user)):
        db = get_db_connection()
        try:
            with db.cursor() as c:
                current = checked_resource_actor(c, actor, load_user, lock=True)
                visit = load(c, visit_id)
                employee(c, current, visit['employee_id'])
                c.execute('SELECT * FROM medical_visit_versions WHERE visit_id=%s AND version=%s', (visit_id, version))
                previous = c.fetchone()
                if not previous: raise HTTPException(404, 'Version introuvable.')
                content = read_pdf({**previous, 'employee_id': visit['employee_id']})
                audit(c, current, visit_id, 'read_version', version=version)
            db.commit()
            return Response(content, media_type='application/pdf', headers={
                'Content-Disposition': "attachment; filename*=UTF-8''" + quote(previous['original_filename'], safe=''),
                'X-Content-Type-Options': 'nosniff'})
        except Exception:
            db.rollback(); raise
        finally:
            db.close()

    @router.put('/medical-visits/{visit_id}')
    def correct(visit_id: int, payload: VisitCorrection, actor=Depends(get_current_user)):
        if payload.visit_date > paris_today():
            raise HTTPException(400, 'La date de visite ne peut pas être future.')
        db, path = get_db_connection(), None
        try:
            with db.cursor() as c:
                current = checked_resource_actor(c, actor, load_user, lock=True)
                visit = load(c, visit_id, lock=True)
                employee(c, current, visit['employee_id'], write=True)
                version_matches(visit, payload)
                if visit['is_archived']:
                    raise HTTPException(409, 'Restaurez cette visite avant modification.')
                if not payload.pdf and str(visit['visit_date']) == payload.visit_date.isoformat():
                    return {'status': 'success'}
                c.execute('''INSERT INTO medical_visit_versions
                    (visit_id,version,visit_date,original_filename,stored_filename,size_bytes,changed_by)
                    VALUES(%s,%s,%s,%s,%s,%s,%s)''',
                    (visit_id, visit['version'], visit['visit_date'], visit['original_filename'],
                     visit['stored_filename'], visit['size_bytes'], current['username']))
                c.execute('UPDATE medical_visits SET visit_date=%s,version=version+%s WHERE id=%s',
                          (payload.visit_date, 0 if payload.pdf else 1, visit_id))
                if payload.pdf:
                    path = store_pdf(c, current, visit, payload.pdf, replace=True)
                audit(c, current, visit_id, 'correct', previous_version=visit['version'],
                      visit_date=payload.visit_date.isoformat())
            db.commit()
            return {'status': 'success'}
        except Exception:
            db.rollback()
            if path: path.unlink(missing_ok=True)
            raise
        finally:
            db.close()

    @router.patch('/medical-visits/{visit_id}/archive')
    def archive(visit_id: int, payload: VisitArchive, actor=Depends(get_current_user)):
        db = get_db_connection()
        try:
            with db.cursor() as c:
                current = checked_resource_actor(c, actor, load_user, lock=True)
                if not is_admin(current): raise HTTPException(403, 'Archivage réservé aux administrateurs.')
                visit = load(c, visit_id, lock=True)
                employee(c, current, visit['employee_id'])
                version_matches(visit, payload)
                if bool(visit['is_archived']) != payload.is_archived:
                    c.execute('UPDATE medical_visits SET is_archived=%s,version=version+1 WHERE id=%s', (int(payload.is_archived), visit_id))
                    audit(c, current, visit_id, 'archive' if payload.is_archived else 'restore')
            db.commit()
            return {'status': 'success'}
        except Exception:
            db.rollback(); raise
        finally:
            db.close()

    @router.delete('/medical-visits/{visit_id}')
    def remove(visit_id: int, payload: VisitDelete, actor=Depends(get_current_user)):
        db, paths = get_db_connection(), []
        try:
            with db.cursor() as c:
                current = checked_resource_actor(c, actor, load_user, lock=True)
                if not is_super_admin(current): raise HTTPException(403, 'Suppression réservée au Super Admin.')
                visit = load(c, visit_id, lock=True)
                employee(c, current, visit['employee_id'])
                version_matches(visit, payload)
                if not visit['is_archived']: raise HTTPException(409, 'Archivez cette visite avant sa suppression.')
                if payload.confirmation != DELETE_WORD: raise HTTPException(400, 'Confirmation incorrecte.')
                c.execute('SELECT stored_filename FROM medical_visit_versions WHERE visit_id=%s', (visit_id,))
                names = {row['stored_filename'] for row in c.fetchall()} | {visit['stored_filename']}
                folder = medical_storage_dir().resolve()
                for name in filter(None, names):
                    if Path(name).name != name or len(name) != 68:
                        raise HTTPException(409, 'Référence de fichier invalide. Suppression interrompue.')
                    path = (folder / name).resolve()
                    if path.parent != folder: raise HTTPException(409, 'Référence de fichier invalide.')
                    paths.append(path)
                c.execute('DELETE FROM medical_visit_versions WHERE visit_id=%s', (visit_id,))
                c.execute('DELETE FROM medical_visit_events WHERE visit_id=%s', (visit_id,))
                c.execute('DELETE FROM medical_visits WHERE id=%s', (visit_id,))
                append_deletion_event(c, current, 'medical_visit', visit_id,
                                      {'deleted_records': 1, 'deleted_files': len(paths), 'previous_is_archived': True})
            db.commit()
        except Exception:
            db.rollback(); raise
        finally:
            db.close()
        return {'status': 'success', 'cleanup_pending': cleanup_files(paths)}
