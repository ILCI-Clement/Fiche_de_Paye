"""Reversible account access changes; historical business records remain untouched."""
from fastapi import Depends, HTTPException
from pydantic import BaseModel, StrictBool


class AccountActivation(BaseModel):
    is_active: StrictBool


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
    cursor.execute('''CREATE TABLE IF NOT EXISTS account_status_events (
        id BIGINT AUTO_INCREMENT PRIMARY KEY,
        target_username VARCHAR(255) NOT NULL, actor_username VARCHAR(255) NOT NULL,
        was_active TINYINT(1) NOT NULL, is_active TINYINT(1) NOT NULL,
        created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
        KEY account_status_target (target_username,id))''')


def register_account_lifecycle(app, *, get_db_connection, get_current_user, load_user, public_user):
    @app.patch('/users/{username}/activation')
    def activate(username: str, payload: AccountActivation, actor=Depends(get_current_user)):
        db = get_db_connection()
        try:
            with db.cursor() as cursor:
                # Same lock order as role updates, including a fresh actor check after waiting.
                cursor.execute('SELECT username FROM users ORDER BY username FOR UPDATE')
                cursor.fetchall()
                current = load_user(cursor, actor['username'])
                if (not current or not current.get('is_active', True)
                    or 'Admin' not in current['role_tags']
                    or current.get('session_version', 0) != actor.get('session_version', 0)):
                    raise HTTPException(403, 'Seul un administrateur actif peut modifier l’accès au compte.')
                target = load_user(cursor, username)
                if not target:
                    raise HTTPException(404, 'Compte introuvable.')
                if target['username'] == current['username'] and not payload.is_active:
                    raise HTTPException(400, 'Vous ne pouvez pas désactiver votre propre compte.')
                was_active = bool(target.get('is_active', True))
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

    @app.get('/users/{username}/activation-history')
    def history(username: str, actor=Depends(get_current_user)):
        if 'Admin' not in actor['role_tags']:
            raise HTTPException(403, 'Historique réservé aux administrateurs.')
        db = get_db_connection()
        try:
            with db.cursor() as cursor:
                cursor.execute('''SELECT actor_username,was_active,is_active,created_at FROM account_status_events
                    WHERE target_username=%s ORDER BY id DESC LIMIT 100''', (username,))
                return {'events': cursor.fetchall()}
        finally:
            db.close()
