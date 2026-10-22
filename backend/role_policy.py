"""Shared account roles and privileged resource guards.

SuperAdmin inherits Admin capabilities. Persisted role changes must still be checked
against the fresh actor and target; role inheritance never grants delegation rights.
"""
import json
import re

from fastapi import HTTPException


VALID_ROLES = frozenset({'SuperAdmin', 'Admin', 'Responsable', 'Employe'})
ROLE_PRIORITY = ('SuperAdmin', 'Admin', 'Responsable', 'Employe')


def canonical_role_tags(tags):
    if not isinstance(tags, (list, tuple, set, frozenset)):
        return []
    recognized = {tag for tag in tags if isinstance(tag, str) and tag in VALID_ROLES}
    if 'SuperAdmin' in recognized:
        recognized.add('Admin')
    return [role for role in ROLE_PRIORITY if role in recognized]


def _roles(user):
    raw = user.get('role_tags')
    if isinstance(raw, str):
        try:
            raw = json.loads(raw)
        except (ValueError, TypeError):
            raw = None
    tags = canonical_role_tags(raw)
    if tags:
        return tags
    if user.get('role') in VALID_ROLES:
        return canonical_role_tags([user['role']])
    return ['Admin'] if user.get('is_admin') else []


def is_super_admin(user):
    return bool(user and 'SuperAdmin' in _roles(user))


def is_admin(user):
    return bool(user and 'Admin' in _roles(user))


def can_manage_account(actor, target):
    return is_admin(actor) and (is_super_admin(actor) or not is_super_admin(target))


def can_assign_roles(actor, requested_tags, target=None):
    if (not isinstance(requested_tags, (list, tuple, set, frozenset))
            or not requested_tags or any(not isinstance(tag, str) or tag not in VALID_ROLES for tag in requested_tags)
            or not is_admin(actor)):
        return False
    return is_super_admin(actor) or ('SuperAdmin' not in requested_tags and not is_super_admin(target))


def checked_resource_actor(cursor, actor, load_user, *, lock=False):
    """Lock users first when writing, then re-read account state, version and roles."""
    if lock:
        cursor.execute('SELECT username FROM users ORDER BY username FOR UPDATE')
        cursor.fetchall()
    current = load_user(cursor, actor.get('username', ''))
    if (not current or not current.get('is_active', True) or current.get('login_deleted_at')
            or current.get('session_version', 0) != actor.get('session_version', 0)):
        raise HTTPException(401, 'Session invalide ou expirée. Reconnectez-vous.')
    return current


def ensure_role_audit_schema(cursor):
    # Independent from users: deleting an empty account cannot erase who changed its roles.
    cursor.execute('''CREATE TABLE IF NOT EXISTS account_role_events (
        id BIGINT NOT NULL AUTO_INCREMENT PRIMARY KEY,
        target_username VARCHAR(255) NOT NULL,
        actor_username VARCHAR(255) NOT NULL,
        action VARCHAR(32) NOT NULL,
        roles_before TEXT NOT NULL,
        roles_after TEXT NOT NULL,
        created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
        KEY account_role_target (target_username,id)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4''')
    cursor.execute('''SELECT CHARACTER_SET_NAME AS charset_name,COLLATION_NAME AS collation_name
        FROM information_schema.columns WHERE table_schema=DATABASE()
        AND table_name='users' AND column_name='username' ''')
    metadata = cursor.fetchone()
    if not metadata or not all(re.fullmatch(r'[A-Za-z0-9_]+', metadata.get(key) or '')
                               for key in ('charset_name', 'collation_name')):
        raise RuntimeError('Cannot identify the account username collation.')
    account_text = f"CHARACTER SET {metadata['charset_name']} COLLATE {metadata['collation_name']}"
    cursor.execute(f'''CREATE TABLE IF NOT EXISTS account_username_reservations (
        username VARCHAR(255) {account_text} NOT NULL PRIMARY KEY,
        created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4''')


def record_role_change(cursor, actor_username, target_username, before, after, *, action='roles_changed'):
    cursor.execute('''INSERT INTO account_role_events
        (target_username,actor_username,action,roles_before,roles_after) VALUES (%s,%s,%s,%s,%s)''',
        (target_username, actor_username, action,
         json.dumps(canonical_role_tags(before)), json.dumps(canonical_role_tags(after))))


def active_super_admin_count(cursor, *, exclude_username=None):
    cursor.execute('''SELECT username,role,role_tags,is_admin,is_active,login_deleted_at
        FROM users ORDER BY username''')
    return sum(1 for user in cursor.fetchall()
               if user['username'] != exclude_username and user.get('is_active', True)
               and not user.get('login_deleted_at') and is_super_admin(user))


def username_is_reserved(cursor, username):
    cursor.execute('SELECT username FROM account_username_reservations WHERE username=%s', (username,))
    return cursor.fetchone() is not None


def reserve_deleted_account_name(cursor, target):
    username = target['username'] if isinstance(target, dict) else target
    cursor.execute('INSERT IGNORE INTO account_username_reservations (username) VALUES (%s)', (username,))
