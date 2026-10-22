"""Public person labels from the explicitly linked register, never account IDs.

Only the register's two name/link columns are read. Callers may reuse a batch
mapping during one request; there is deliberately no process-wide name cache.
"""
from collections.abc import Iterable, Mapping
import re
import sqlite3

import pymysql


UNNAMED_PERSON = 'Nom non renseigné'
_BATCH_SIZE = 500


def _text(value) -> str:
    return value.strip() if isinstance(value, str) else ''


def _usernames(values: Iterable[str]) -> list[str]:
    return sorted({value for value in values if isinstance(value, str) and value})


def _sqlite(cursor) -> bool:
    return isinstance(cursor, sqlite3.Cursor) or isinstance(getattr(cursor, 'connection', None), sqlite3.Connection)


def _rows(cursor) -> list[dict]:
    columns = [column[0] for column in cursor.description or []]
    return [dict(row) if isinstance(row, Mapping) or isinstance(row, sqlite3.Row) else dict(zip(columns, row))
            for row in cursor.fetchall()]


def load_person_names(cursor, usernames: Iterable[str] | None = None) -> dict[str, str]:
    """Read the authoritative full names, safely supporting older installations.

    A missing register is expected during startup and on the legacy server. Other
    database failures are not swallowed. Matching uses explicit linked usernames;
    e-mail addresses and similar-looking names never establish a relationship.
    """
    selected = None if usernames is None else _usernames(usernames)
    if selected == []:
        return {}
    sqlite = _sqlite(cursor)
    if sqlite:
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='personnel_register'")
    else:
        cursor.execute("SHOW TABLES LIKE 'personnel_register'")
    if not cursor.fetchone():
        return {}
    names: dict[str, str] = {}
    batches = [None] if selected is None else [selected[i:i + _BATCH_SIZE] for i in range(0, len(selected), _BATCH_SIZE)]
    try:
        for batch in batches:
            sql = 'SELECT linked_username,full_name FROM personnel_register WHERE linked_username IS NOT NULL'
            if batch is None:
                cursor.execute(sql)
            else:
                placeholders = ','.join(['?' if sqlite else '%s'] * len(batch))
                cursor.execute(f'{sql} AND linked_username IN ({placeholders})', tuple(batch))
            for row in _rows(cursor):
                username = row.get('linked_username')
                name = _text(row.get('full_name'))
                if isinstance(username, str) and username and name:
                    names[username] = name
    except pymysql.err.ProgrammingError as exc:
        # A migration/deletion race can remove the optional table after SHOW.
        if exc.args and exc.args[0] == 1146:
            return {}
        raise
    return names


def person_display_name(account: Mapping | None, names: Mapping[str, str] | None = None) -> str:
    """Return linked full name, then e-mail local part, then a neutral label."""
    account = account or {}
    full_name = _text(names.get(account.get('username'))) if names is not None else _text(account.get('full_name'))
    if full_name:
        return full_name
    email = _text(account.get('email'))
    return email.partition('@')[0] if re.fullmatch(r'[^@\s]+@[^@\s]+', email) else UNNAMED_PERSON


def annotate_person_names(accounts: Iterable[dict], names: Mapping[str, str]) -> list[dict]:
    """Annotate already-authorized account DTOs in place without copying identities."""
    result = list(accounts)
    for account in result:
        account['full_name'] = _text(names.get(account.get('username'))) or None
        account['name'] = account['display_name'] = person_display_name(account, names)
    return result


def load_person_accounts(cursor, usernames: Iterable[str] | None = None, *, person_names=None) -> dict[str, dict]:
    """Load a public name directory without any credential or sensitive fields."""
    selected = None if usernames is None else _usernames(usernames)
    if selected == []:
        return {}
    if selected is None:
        cursor.execute('SELECT username,email FROM users ORDER BY username')
        accounts = _rows(cursor)
    else:
        accounts = []
        for offset in range(0, len(selected), _BATCH_SIZE):
            batch = selected[offset:offset + _BATCH_SIZE]
            placeholders = ','.join(['?' if _sqlite(cursor) else '%s'] * len(batch))
            cursor.execute(f'SELECT username,email FROM users WHERE username IN ({placeholders})', tuple(batch))
            accounts.extend(_rows(cursor))
    names = load_person_names(cursor, [account['username'] for account in accounts]) if person_names is None else person_names
    return {account['username']: account for account in annotate_person_names(accounts, names)}


def load_display_names(cursor, usernames: Iterable[str]) -> dict[str, str]:
    """Resolve historical actor/owner IDs with a bounded batch of public fields."""
    selected = _usernames(usernames)
    if not selected:
        return {}
    return {username: account['display_name'] for username, account in load_person_accounts(cursor, selected).items()}
