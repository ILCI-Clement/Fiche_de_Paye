"""Atomic administrator edits of direct reports; no parallel organization model."""
from fastapi import Depends, HTTPException
from pydantic import BaseModel, ConfigDict, Field, StrictBool

from .person_visibility import account_list_visible, archived_usernames, daily_person
from .role_policy import checked_resource_actor, is_admin, can_manage_account


def ensure_management_assignment_schema(cursor):
    cursor.execute('''CREATE TABLE IF NOT EXISTS management_assignment_events (
        id BIGINT NOT NULL AUTO_INCREMENT PRIMARY KEY,
        actor_username VARCHAR(255) NOT NULL, employee_username VARCHAR(255) NOT NULL,
        before_manager VARCHAR(255) NULL, after_manager VARCHAR(255) NULL,
        created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
        KEY assignment_employee (employee_username,id)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4''')


class ReportChange(BaseModel):
    model_config = ConfigDict(extra='forbid')
    username: str = Field(min_length=1, max_length=255)
    expected_manager_id: str | None = Field(max_length=255)
    assigned: StrictBool


class ReportsUpdate(BaseModel):
    model_config = ConfigDict(extra='forbid')
    changes: list[ReportChange] = Field(min_length=1, max_length=100)


def register_report_routes(app, *, get_db_connection, get_current_user, load_user, valid_manager, would_cycle):
    @app.put('/users/{manager_username}/reports')
    def update_reports(manager_username: str, payload: ReportsUpdate, actor=Depends(get_current_user)):
        db = get_db_connection()
        try:
            with db.cursor() as cursor:
                actor = checked_resource_actor(cursor, actor, load_user, lock=True)
                if not is_admin(actor):
                    raise HTTPException(403, 'Accès réservé aux administrateurs.')
                manager = load_user(cursor, manager_username)
                if (not manager or not account_list_visible(manager, actor)
                        or not can_manage_account(actor, manager)):
                    raise HTTPException(404, 'Compte introuvable ou inaccessible.')
                if any(change.assigned for change in payload.changes) and (
                        not valid_manager(manager) or manager.get('login_deleted_at')):
                    raise HTTPException(400, 'Choisissez un compte Responsable ou Admin actif.')
                names = [change.username for change in payload.changes]
                if len(names) != len(set(names)):
                    raise HTTPException(400, 'Une personne ne peut apparaître qu’une fois.')
                archived = archived_usernames(cursor)
                if manager_username in archived and any(change.assigned for change in payload.changes):
                    raise HTTPException(409, 'Ce responsable est archivé.')
                edits = []
                for change in payload.changes:
                    target = load_user(cursor, change.username)
                    if (not target or not account_list_visible(target, actor)
                            or not can_manage_account(actor, target)):
                        raise HTTPException(404, 'Personne introuvable ou inaccessible.')
                    if change.username == manager_username:
                        raise HTTPException(400, 'Un compte ne peut pas être son propre responsable.')
                    before = target.get('manager_username') or None
                    if before != change.expected_manager_id:
                        raise HTTPException(409, 'Une affectation a changé. Actualisez la liste avant de réessayer.')
                    if not change.assigned and before != manager_username:
                        raise HTTPException(409, 'Cette personne ne dépend plus de ce responsable.')
                    if change.assigned and (not daily_person(target) or change.username in archived):
                        raise HTTPException(409, 'Une personne sélectionnée est désactivée ou archivée.')
                    after = manager_username if change.assigned else None
                    if before != after:
                        edits.append((change.username, before, after))
                # Validate cycles against the final graph, not a partially applied list.
                for username, _, after in edits:
                    cursor.execute('UPDATE users SET manager_username=%s WHERE username=%s', (after, username))
                for username, before, after in edits:
                    if after and would_cycle(cursor, username, after):
                        raise HTTPException(400, 'Cette affectation créerait une boucle hiérarchique.')
                    cursor.execute('''INSERT INTO management_assignment_events
                        (actor_username,employee_username,before_manager,after_manager) VALUES(%s,%s,%s,%s)''',
                        (actor['username'], username, before, after))
            db.commit()
            return {'status': 'success', 'updated': len(edits)}
        except Exception:
            db.rollback()
            raise
        finally:
            db.close()
