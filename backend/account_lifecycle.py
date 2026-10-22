"""Reversible account access changes; historical business records remain untouched."""
from fastapi import Depends, HTTPException
from pydantic import BaseModel, StrictBool
import secrets

if __package__:
    from .person_names import load_display_names
    from .role_policy import (active_super_admin_count, can_manage_account, checked_resource_actor,
                              is_admin, is_super_admin)
else:
    from person_names import load_display_names
    from role_policy import (active_super_admin_count, can_manage_account, checked_resource_actor,
                             is_admin, is_super_admin)


class AccountActivation(BaseModel):
    is_active: StrictBool


class LoginDeletion(BaseModel):
    confirmation: str


def ensure_account_lifecycle_schema(cursor):
    # Historical snapshots have a zero-date default on reset_token_expires.
    # Relax only zero-date checks for this connection's additive DDL, then restore.
    cursor.execute('SELECT @@SESSION.sql_mode AS value')
    previous = cursor.fetchone()['value']
    mode = ','.join(part for part in previous.split(',') if part not in {'NO_ZERO_DATE', 'NO_ZERO_IN_DATE'})
    cursor.execute('SET SESSION sql_mode=%s', (mode,))
    try:
        _add_account_columns(cursor)
    finally:
        cursor.execute('SET SESSION sql_mode=%s', (previous,))


def _add_account_columns(cursor):
    cursor.execute('SHOW COLUMNS FROM users')
    columns = {row['Field'] for row in cursor.fetchall()}
    if 'is_active' not in columns:
        cursor.execute('ALTER TABLE users ADD COLUMN is_active TINYINT(1) NOT NULL DEFAULT 1')
    if 'session_version' not in columns:
        cursor.execute('ALTER TABLE users ADD COLUMN session_version INT NOT NULL DEFAULT 0')
    if 'login_deleted_at' not in columns:
        cursor.execute('ALTER TABLE users ADD COLUMN login_deleted_at DATETIME NULL DEFAULT NULL')
    cursor.execute('''CREATE TABLE IF NOT EXISTS account_status_events (
        id BIGINT AUTO_INCREMENT PRIMARY KEY,
        target_username VARCHAR(255) NOT NULL, actor_username VARCHAR(255) NOT NULL,
        was_active TINYINT(1) NOT NULL, is_active TINYINT(1) NOT NULL,
        created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
        KEY account_status_target (target_username,id))''')
    cursor.execute('SHOW COLUMNS FROM account_status_events')
    if 'action' not in {row['Field'] for row in cursor.fetchall()}:
        cursor.execute("ALTER TABLE account_status_events ADD COLUMN action VARCHAR(20) NOT NULL DEFAULT 'activation'")


def register_account_lifecycle(app, *, get_db_connection, get_current_user, load_user, public_user, hash_password):
    def locked_target(cursor, username, actor):
        current = checked_resource_actor(cursor, actor, load_user, lock=True)
        if not is_admin(current):
            raise HTTPException(403, 'Seul un administrateur actif peut modifier l’accès au compte.')
        target = load_user(cursor, username)
        if not target:
            raise HTTPException(404, 'Compte introuvable.')
        if not can_manage_account(current, target):
            raise HTTPException(403, 'Seul un Super-administrateur peut modifier ce compte.')
        return current, target

    @app.patch('/users/{username}/activation')
    def activate(username: str, payload: AccountActivation, actor=Depends(get_current_user)):
        db = get_db_connection()
        try:
            with db.cursor() as cursor:
                # Same lock order as role updates, including a fresh actor check after waiting.
                current, target = locked_target(cursor, username, actor)
                if target.get('login_deleted_at'):
                    raise HTTPException(409, 'L’accès de connexion a été supprimé et ne peut pas être réactivé.')
                if target['username'] == current['username'] and not payload.is_active:
                    raise HTTPException(400, 'Vous ne pouvez pas désactiver votre propre compte.')
                was_active = bool(target.get('is_active', True))
                if was_active and not payload.is_active and is_super_admin(target):
                    if active_super_admin_count(cursor, exclude_username=target['username']) == 0:
                        raise HTTPException(409, 'Le dernier Super-administrateur actif doit être conservé.')
                if was_active != payload.is_active:
                    cursor.execute('''UPDATE users SET is_active=%s, session_version=session_version+1,
                        remember_token_hash=NULL,remember_token_expires=NULL,reset_token=NULL,reset_token_expires=CURRENT_TIMESTAMP
                        WHERE username=%s''', (payload.is_active, target['username']))
                    cursor.execute('''INSERT INTO account_status_events
                        (target_username,actor_username,was_active,is_active) VALUES (%s,%s,%s,%s)''',
                        (target['username'], current['username'], was_active, payload.is_active))
                updated = load_user(cursor, target['username'])
            db.commit()
            return {'status': 'success', 'user': public_user(updated)}
        except Exception:
            db.rollback()
            raise
        finally:
            db.close()

    @app.delete('/users/{username}/login')
    def delete_login(username: str, payload: LoginDeletion, actor=Depends(get_current_user)):
        db = get_db_connection()
        try:
            with db.cursor() as cursor:
                current, target = locked_target(cursor, username, actor)
                if not is_super_admin(current):
                    raise HTTPException(403, 'Seul un Super-administrateur peut supprimer un accès de connexion.')
                if target['username'] == current['username']:
                    raise HTTPException(400, 'Vous ne pouvez pas supprimer votre propre accès.')
                confirmations = {target['username']}
                email = str(target.get('email') or '').strip()
                if email:
                    confirmations.add(email)
                if payload.confirmation not in confirmations:
                    raise HTTPException(400, 'Saisissez l’adresse e-mail exacte du compte pour confirmer.')
                if target.get('is_active', True):
                    raise HTTPException(409, 'Désactivez le compte avant de supprimer son accès de connexion.')
                if not target.get('login_deleted_at'):
                    # Keep the identity row and every business relation as an archive anchor.
                    # Replace the credential; no physical deletion or file unlink is performed.
                    cursor.execute('''UPDATE users SET login_deleted_at=CURRENT_TIMESTAMP,
                        is_active=0,session_version=session_version+1,password_hash=%s,
                        remember_token_hash=NULL,remember_token_expires=NULL,
                        reset_token=NULL,reset_token_expires=CURRENT_TIMESTAMP WHERE username=%s''',
                        (hash_password(secrets.token_urlsafe(32)), target['username']))
                    cursor.execute('''INSERT INTO account_status_events
                        (target_username,actor_username,was_active,is_active,action) VALUES (%s,%s,0,0,'login_deleted')''',
                        (target['username'], current['username']))
                updated = load_user(cursor, target['username'])
            db.commit()
            return {'status': 'success', 'user': public_user(updated)}
        except Exception:
            db.rollback()
            raise
        finally:
            db.close()

    @app.get('/users/{username}/activation-history')
    def history(username: str, actor=Depends(get_current_user)):
        db = get_db_connection()
        try:
            with db.cursor() as cursor:
                current = checked_resource_actor(cursor, actor, load_user)
                if not is_admin(current):
                    raise HTTPException(403, 'Historique réservé aux administrateurs.')
                cursor.execute('''SELECT actor_username,was_active,is_active,action,created_at FROM account_status_events
                    WHERE target_username=%s ORDER BY id DESC LIMIT 100''', (username,))
                events = cursor.fetchall()
                names = load_display_names(cursor, [event['actor_username'] for event in events])
                for event in events:
                    event['actor_display_name'] = names.get(event['actor_username'], 'Ancien compte')
                return {'events': events}
        finally:
            db.close()
