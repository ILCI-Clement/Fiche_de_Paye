"""Shared, database-backed password failure cooldown for every API worker."""
from datetime import timedelta
from math import ceil

from fastapi import HTTPException

FAILURE_LIMIT = 5
COOLDOWN_SECONDS = 10 * 60
FAILURE_RESET_SECONDS = 10 * 60


def ensure_login_guard_schema(cursor, columns):
    additions = []
    if 'login_failed_attempts' not in columns:
        additions.append('ADD COLUMN login_failed_attempts INT NOT NULL DEFAULT 0')
    if 'login_locked_until' not in columns:
        # UTC without session time-zone conversion; compare against MySQL UTC_TIMESTAMP.
        additions.append('ADD COLUMN login_locked_until DATETIME(6) NULL')
    if 'login_last_failed_at' not in columns:
        additions.append('ADD COLUMN login_last_failed_at DATETIME(6) NULL')
    if not additions:
        return
    # Legacy users tables can still declare zero-date reset-token defaults.
    # Validate this additive ALTER without rewriting those old fields or changing
    # global SQL mode. Restore this connection's strict mode before other work.
    cursor.execute('SELECT @@SESSION.sql_mode AS sql_mode')
    mode = cursor.fetchone()['sql_mode']
    compatible = ','.join(part for part in mode.split(',') if part not in {'NO_ZERO_DATE', 'NO_ZERO_IN_DATE'})
    try:
        cursor.execute('SET SESSION sql_mode=%s', (compatible,))
        cursor.execute('ALTER TABLE users ' + ','.join(additions))
    finally:
        cursor.execute('SET SESSION sql_mode=%s', (mode,))


def lock_login_attempt(cursor, username, identifier):
    """Serialize attempts for one canonical account, including username/email aliases."""
    cursor.execute('''SELECT login_failed_attempts,login_locked_until,login_last_failed_at,UTC_TIMESTAMP(6) AS login_now
        FROM users WHERE username=%s AND
        (LOWER(TRIM(username))=LOWER(%s) OR LOWER(TRIM(email))=LOWER(%s)) FOR UPDATE''',
                   (username, identifier, identifier))
    return cursor.fetchone()


def cooldown_error(seconds):
    return HTTPException(429, 'Trop de tentatives de connexion. Réessayez après le délai de sécurité.',
                         headers={'Retry-After': str(max(1, ceil(seconds)))})


def check_cooldown(state):
    until = state['login_locked_until']
    if until is not None and until > state['login_now']:
        raise cooldown_error((until - state['login_now']).total_seconds())


def record_failure(cursor, username, state):
    # Expiry or a ten-minute gap starts a new sequence. No background job is
    # required: the next serialized attempt applies the reset before counting.
    # Missing timestamps from the earlier schema also start a fresh sequence.
    now = state['login_now']
    last_failure = state['login_last_failed_at']
    expired = state['login_locked_until'] is not None and state['login_locked_until'] <= now
    idle = last_failure is None or now - last_failure >= timedelta(seconds=FAILURE_RESET_SECONDS)
    failures = 1 if expired or idle else min(FAILURE_LIMIT, int(state['login_failed_attempts']) + 1)
    locked = failures >= FAILURE_LIMIT
    until = state['login_now'] + timedelta(seconds=COOLDOWN_SECONDS) if locked else None
    cursor.execute('UPDATE users SET login_failed_attempts=%s,login_locked_until=%s,login_last_failed_at=%s WHERE username=%s',
                   (failures, until, now, username))
    return locked


def reset_failures(cursor, username):
    cursor.execute('UPDATE users SET login_failed_attempts=0,login_locked_until=NULL,login_last_failed_at=NULL WHERE username=%s',
                   (username,))
