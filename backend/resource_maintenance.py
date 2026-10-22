"""Archive controls and guarded Super Admin deletion for administrative resources."""
from typing import Literal
from fastapi import Depends, HTTPException, Query
from pydantic import BaseModel, Field, StrictBool
from .destructive_audit import append_deletion_event, append_archive_event, ensure_destructive_audit_schema
from .role_policy import is_admin, is_super_admin
from .person_names import load_display_names


ResourceKind = Literal['group', 'leave_request', 'annual_interview']
TABLES = {'group': 'organization_groups', 'leave_request': 'leave_requests',
          'annual_interview': 'annual_interviews'}


class ResourceArchive(BaseModel):
    is_archived: StrictBool


class ResourceDeletion(BaseModel):
    confirmation: str = Field(max_length=255)


def register_resource_maintenance(app, *, get_db_connection, get_current_user, load_user):
    @app.on_event('startup')
    def migrate_resource_maintenance():
        db = get_db_connection()
        try:
            with db.cursor() as cursor:
                ensure_destructive_audit_schema(cursor)
                for table in ('leave_requests', 'annual_interviews'):
                    cursor.execute(f'SHOW COLUMNS FROM {table}')
                    if 'is_archived' not in {row['Field'] for row in cursor.fetchall()}:
                        cursor.execute(f'ALTER TABLE {table} ADD COLUMN is_archived TINYINT(1) NOT NULL DEFAULT 0')
            db.commit()
        finally:
            db.close()

    def fresh_admin(cursor, actor, *, write=False, deletion=False):
        if write:
            cursor.execute('SELECT username FROM users ORDER BY username FOR UPDATE')
            cursor.fetchall()
        current = load_user(cursor, actor['username'])
        if (not current or not current.get('is_active', True) or current.get('login_deleted_at')
                or current.get('session_version', 0) != actor.get('session_version', 0)
                or not is_admin(current) or (deletion and not is_super_admin(current))):
            raise HTTPException(403, 'Autorisation administrative insuffisante ou session expirée.')
        return current

    @app.get('/maintenance/resources/{kind}')
    def resources(kind: ResourceKind, archive: Literal['active', 'archived'] = 'active',
                  page: int = Query(default=1, ge=1, le=1_000_000),
                  page_size: int = Query(default=25, ge=1, le=100), actor=Depends(get_current_user)):
        db = get_db_connection()
        try:
            with db.cursor() as cursor:
                fresh_admin(cursor, actor)
                table = TABLES[kind]
                predicate = 'is_active=%s' if kind == 'group' else 'is_archived=%s'
                value = archive == 'active' if kind == 'group' else archive == 'archived'
                cursor.execute(f'SELECT COUNT(*) AS total FROM {table} WHERE {predicate}', (value,))
                total = cursor.fetchone()['total']
                columns = {'group': 'id,name,is_active',
                           'leave_request': 'id,employee_username,start_date,end_date,status,is_archived',
                           'annual_interview': 'id,employee_username,interview_year,completed,is_archived'}[kind]
                cursor.execute(f'SELECT {columns} FROM {table} WHERE {predicate} ORDER BY id DESC LIMIT %s OFFSET %s',
                               (value, page_size, (page - 1) * page_size))
                rows = cursor.fetchall()
                names = (load_display_names(cursor, [row.get('employee_username') for row in rows])
                         if kind != 'group' else {})
                items = []
                for row in rows:
                    if kind == 'group':
                        label, detail, archived = row['name'], 'Département', not bool(row['is_active'])
                    elif kind == 'leave_request':
                        label = names.get(row['employee_username'], 'Ancien compte')
                        detail = f"{row['start_date']} – {row['end_date']} · {row['status']}"
                        archived = bool(row['is_archived'])
                    else:
                        label = names.get(row['employee_username'], 'Ancien compte')
                        detail = f"{row['interview_year']} · {'Terminé' if row['completed'] else 'À faire'}"
                        archived = bool(row['is_archived'])
                    items.append({'id': str(row['id']), 'label': label, 'detail': detail, 'is_archived': archived})
                return {'items': items, 'total': total, 'page': page, 'page_size': page_size}
        finally:
            db.close()

    def mutate(kind, resource_id, actor, *, archived=None, confirmation=None):
        db = get_db_connection()
        try:
            with db.cursor() as cursor:
                current = fresh_admin(cursor, actor, write=True, deletion=confirmation is not None)
                if resource_id < 1 or resource_id > 9_007_199_254_740_991:
                    raise HTTPException(404, 'Ressource introuvable.')
                table = TABLES[kind]
                cursor.execute(f'SELECT * FROM {table} WHERE id=%s FOR UPDATE', (resource_id,))
                resource = cursor.fetchone()
                if not resource:
                    raise HTTPException(404, 'Ressource introuvable.')
                old_archived = not bool(resource['is_active']) if kind == 'group' else bool(resource['is_archived'])
                if confirmation is not None:
                    if confirmation != str(resource_id):
                        raise HTTPException(400, 'Saisissez le numéro exact pour confirmer la suppression définitive.')
                    if not old_archived:
                        raise HTTPException(409, 'Archivez la ressource avant sa suppression définitive.')
                    if kind == 'group':
                        cursor.execute('SELECT COUNT(*) AS total FROM user_group_memberships WHERE group_id=%s', (resource_id,))
                        if cursor.fetchone()['total']:
                            raise HTTPException(409, 'Réaffectez les membres et responsables avant de supprimer ce département.')
                    if kind == 'annual_interview':
                        owner = load_user(cursor, resource['employee_username'])
                        if owner and owner.get('is_active', True) and not owner.get('login_deleted_at'):
                            raise HTTPException(409, 'Les entretiens des comptes actifs sont générés automatiquement. Conservez cette archive ou désactivez d’abord le compte.')
                    history_count = 0
                    if kind == 'leave_request':
                        cursor.execute('SELECT COUNT(*) AS total FROM leave_events WHERE request_id=%s', (resource_id,))
                        history_count = int(cursor.fetchone()['total'])
                        cursor.execute('DELETE FROM leave_events WHERE request_id=%s', (resource_id,))
                    cursor.execute(f'DELETE FROM {table} WHERE id=%s', (resource_id,))
                    append_deletion_event(cursor, current, kind, resource_id,
                                          {'deleted_records': 1, 'history_count': history_count, 'previous_is_archived': True})
                elif old_archived != archived:
                    if kind == 'leave_request' and resource['status'] == 'pending' and archived:
                        raise HTTPException(409, 'Traitez ou retirez la demande en attente avant de l’archiver.')
                    column, value = ('is_active', not archived) if kind == 'group' else ('is_archived', archived)
                    cursor.execute(f'UPDATE {table} SET {column}=%s WHERE id=%s', (value, resource_id))
                    append_archive_event(cursor, current, kind, resource_id, archived)
            db.commit()
            return {'status': 'success', 'deleted': confirmation is not None, 'is_archived': archived}
        except Exception:
            db.rollback()
            raise
        finally:
            db.close()

    @app.patch('/maintenance/resources/{kind}/{resource_id}/archive')
    def archive_resource(kind: ResourceKind, resource_id: int, payload: ResourceArchive, actor=Depends(get_current_user)):
        return mutate(kind, resource_id, actor, archived=payload.is_archived)

    @app.delete('/maintenance/resources/{kind}/{resource_id}')
    def delete_resource(kind: ResourceKind, resource_id: int, payload: ResourceDeletion, actor=Depends(get_current_user)):
        return mutate(kind, resource_id, actor, confirmation=payload.confirmation)
