"""Unique live login mailboxes; deleted identities keep their historical address."""
import re
from collections import defaultdict

from fastapi import HTTPException


def ensure_account_email_schema(cursor):
    cursor.execute("SELECT GET_LOCK(CONCAT('rh:login-email:',DATABASE()),30) AS acquired")
    if cursor.fetchone()['acquired'] != 1:
        raise RuntimeError('Login email migration is busy.')
    previous = None
    try:
        # Legacy users tables contain a zero-date default on reset_token_expires.
        # Apply the existing connection-local DDL compatibility, never global modes.
        cursor.execute('SELECT @@SESSION.sql_mode AS value')
        previous = cursor.fetchone()['value']
        mode = ','.join(p for p in previous.split(',') if p not in {'NO_ZERO_DATE', 'NO_ZERO_IN_DATE'})
        cursor.execute('SET SESSION sql_mode=%s', (mode,))
        _ensure_account_email_schema(cursor)
    finally:
        if previous is not None:
            cursor.execute('SET SESSION sql_mode=%s', (previous,))
        cursor.execute("SELECT RELEASE_LOCK(CONCAT('rh:login-email:',DATABASE()))")


def _ensure_account_email_schema(cursor):
    """Install the live-login constraint before removing the old all-history one."""
    cursor.execute("SHOW COLUMNS FROM users LIKE 'login_email_key'")
    column = cursor.fetchone()
    if not column:
        cursor.execute('''ALTER TABLE users ADD COLUMN login_email_key VARCHAR(255)
            GENERATED ALWAYS AS (CASE WHEN login_deleted_at IS NULL
                THEN LOWER(TRIM(email)) ELSE NULL END) STORED''')
    elif 'GENERATED' not in column.get('Extra', '').upper():
        raise RuntimeError('Unexpected login email key definition.')
    cursor.execute('SHOW INDEX FROM users')
    indexes = defaultdict(list)
    for row in cursor.fetchall():
        if not row['Non_unique']:
            indexes[row['Key_name']].append((row['Seq_in_index'], row['Column_name']))
    columns = {name: [c for _, c in sorted(items)] for name, items in indexes.items()}
    if 'users_login_email_unique' not in columns:
        cursor.execute('ALTER TABLE users ADD UNIQUE KEY users_login_email_unique (login_email_key)')
    elif columns['users_login_email_unique'] != ['login_email_key']:
        raise RuntimeError('Unexpected login email uniqueness constraint.')
    for name, fields in columns.items():
        if name != 'PRIMARY' and fields == ['email']:
            if not re.fullmatch(r'[A-Za-z0-9_]+', name):
                raise RuntimeError('Unexpected email index name.')
            cursor.execute(f'ALTER TABLE users DROP INDEX `{name}`')


def require_available_login_email(cursor, email, *, exclude_username=''):
    # The caller locks users in a stable order. The generated unique key also
    # protects against concurrent writers that do not use the application lock.
    cursor.execute('''SELECT username FROM users WHERE LOWER(TRIM(email))=LOWER(%s)
        AND login_deleted_at IS NULL AND username<>%s LIMIT 1''', (email.strip(), exclude_username))
    if cursor.fetchone():
        raise HTTPException(400, 'Cette adresse e-mail est déjà utilisée par un compte. '
                            'Un compte désactivé ou archivé conserve son adresse.')


def change_linked_login_email(cursor, target, email, actor):
    from .role_policy import can_manage_account
    from .person_visibility import account_list_visible
    if not target or not can_manage_account(actor, target) or not account_list_visible(target, actor):
        raise HTTPException(403, 'Vous ne pouvez pas modifier l’e-mail de connexion de ce compte.')
    if target.get('login_deleted_at'):
        raise HTTPException(409, 'L’accès de ce compte a été supprimé. Dissociez la fiche avant de modifier son e-mail.')
    if not email:
        raise HTTPException(400, 'Le compte associé doit conserver une adresse e-mail de connexion.')
    require_available_login_email(cursor, email, exclude_username=target['username'])
    if email != target.get('email'):
        cursor.execute('''UPDATE users SET email=%s,session_version=session_version+1,
            remember_token_hash=NULL,remember_token_expires=NULL,reset_token=NULL,
            reset_token_expires=CURRENT_TIMESTAMP WHERE username=%s''', (email, target['username']))


def sync_registry_email(cursor, username, email, actor_username):
    from .personnel_store import load_employee_detail, append_personnel_event
    cursor.execute('SELECT employee_id FROM personnel_register WHERE linked_username=%s FOR UPDATE', (username,))
    row = cursor.fetchone()
    if not row:
        return
    before = load_employee_detail(cursor, row['employee_id'], lock=True)
    if before['employee']['email'] == email:
        return
    cursor.execute('UPDATE personnel_register SET email=%s,version=version+1 WHERE employee_id=%s',
                   (email, row['employee_id']))
    after = load_employee_detail(cursor, row['employee_id'])
    append_personnel_event(cursor, row['employee_id'], actor_username, 'email_updated', before, after)
