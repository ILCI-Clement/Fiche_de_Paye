"""Public person labels from structured identities, never account IDs.

Only account and register name/link columns are read. Callers may reuse a batch
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


def _missing_name_column(error):
    message = str(error)
    if isinstance(error, sqlite3.OperationalError):
        return bool(re.search(r'no such column: (?:\w+\.)?(first_name|last_name)\b', message))
    return (isinstance(error, pymysql.err.OperationalError) and error.args[0] == 1054
            and bool(re.search(r"Unknown column '[^']*\b(first_name|last_name)'", message)))


def _structured_name(row):
    first, last = _text(row.get('first_name')), _text(row.get('last_name'))
    return f'{first} {last}' if first and last else ''


def load_person_names(cursor, usernames: Iterable[str] | None = None) -> dict[str, str]:
    """Registered structured name, account structured name, then legacy full name.

    Missing new columns support old API installations during a staged release.
    Never split old labels or infer a relationship by name or e-mail.
    """
    selected = None if usernames is None else _usernames(usernames)
    if selected == []:
        return {}
    sqlite = _sqlite(cursor)
    if sqlite:
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='personnel_register'")
    else:
        cursor.execute("SHOW TABLES LIKE 'personnel_register'")
    register_exists = bool(cursor.fetchone())
    names: dict[str, str] = {}
    registered_structured: dict[str, str] = {}
    batches = [None] if selected is None else [selected[i:i + _BATCH_SIZE] for i in range(0, len(selected), _BATCH_SIZE)]
    try:
        for batch in batches:
            if register_exists:
                sql = 'SELECT linked_username,full_name,first_name,last_name FROM personnel_register WHERE linked_username IS NOT NULL'
                if batch is not None:
                    placeholders = ','.join(['?' if sqlite else '%s'] * len(batch))
                    sql += f' AND linked_username IN ({placeholders})'
                parameters = () if batch is None else tuple(batch)
                try:
                    cursor.execute(sql, parameters)
                except (sqlite3.OperationalError, pymysql.err.OperationalError) as error:
                    if not _missing_name_column(error):
                        raise
                    cursor.execute(sql.replace(',first_name,last_name', ''), parameters)
                for row in _rows(cursor):
                    username = row.get('linked_username')
                    if isinstance(username, str) and username:
                        if _text(row.get('full_name')):
                            names[username] = _text(row['full_name'])
                        if _structured_name(row):
                            registered_structured[username] = _structured_name(row)
    except pymysql.err.ProgrammingError as exc:
        # A migration/deletion race can remove the optional table after SHOW.
        if exc.args and exc.args[0] == 1146:
            return {}
        raise
    for batch in batches:
        sql = 'SELECT username,first_name,last_name FROM users'
        if batch is not None:
            placeholders = ','.join(['?' if sqlite else '%s'] * len(batch))
            sql += f' WHERE username IN ({placeholders})'
        try:
            cursor.execute(sql, () if batch is None else tuple(batch))
        except (sqlite3.OperationalError, pymysql.err.OperationalError) as error:
            if _missing_name_column(error):
                break
            raise
        for row in _rows(cursor):
            if _structured_name(row):
                names[row['username']] = _structured_name(row)
    return {**names, **registered_structured}


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
