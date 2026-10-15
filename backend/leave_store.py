"""Leave persistence. MariaDB in production; SQLite only for the isolated preview."""
import sqlite3
import re
from datetime import datetime, timezone

if __package__:
    from .leave_rules import display_name
else:
    from leave_rules import display_name


def timestamp():
    return datetime.now(timezone.utc).isoformat(timespec='microseconds')


def iso(value):
    return value.isoformat() if hasattr(value, 'isoformat') else value


class LeaveStore:
    def __init__(self, connect, *, write=False):
        self.connect = connect
        self.write = write

    def __enter__(self):
        self.db = self.connect()
        self.sqlite = isinstance(self.db, sqlite3.Connection)
        try:
            self.cursor = self.db.cursor()
            if self.write:
                if self.sqlite:
                    self.db.execute('BEGIN IMMEDIATE')
                else:
                    self.db.begin()
        except Exception:
            self.db.close()
            raise
        return self

    def __exit__(self, kind, value, traceback):
        try:
            if self.write:
                self.db.rollback() if kind else self.db.commit()
        finally:
            self.cursor.close()
            self.db.close()

    def execute(self, sql, params=()):
        return self.cursor.execute(sql.replace('%s', '?') if self.sqlite else sql, params)

    def one(self):
        row = self.cursor.fetchone()
        return dict(row) if row is not None else None

    def many(self):
        return [dict(row) for row in self.cursor.fetchall()]

    def user(self, username):
        self.execute('SELECT username,email,manager_username FROM users WHERE username=%s', (username,))
        return self.one()

    def request(self, request_id, *, lock=False):
        suffix = ' FOR UPDATE' if lock and not self.sqlite else ''
        self.execute('''SELECT r.*, u.manager_username, u.email AS current_email,
            m.email AS manager_email, m.username AS current_manager
            FROM leave_requests r LEFT JOIN users u ON u.username=r.employee_username
            LEFT JOIN users m ON m.username=u.manager_username WHERE r.id=%s''' + suffix, (request_id,))
        return self.one()

    def event(self, request_id, actor, action, previous, status, comment=''):
        self.execute('''INSERT INTO leave_events
            (request_id,actor_username,actor_name,action,previous_status,new_status,comment,created_at)
            VALUES (%s,%s,%s,%s,%s,%s,%s,%s)''',
            (request_id, actor['username'], display_name(actor), action, previous, status, comment, timestamp()))


def ensure_leave_schema(connect):
    with LeaveStore(connect, write=True) as store:
        identity = 'INTEGER PRIMARY KEY AUTOINCREMENT' if store.sqlite else 'BIGINT AUTO_INCREMENT PRIMARY KEY'
        indexes = '' if store.sqlite else ', KEY leave_owner (employee_username,status,id), KEY leave_status (status,id)'
        store.execute(f'''CREATE TABLE IF NOT EXISTS leave_requests (
            id {identity}, employee_username VARCHAR(255) NULL,
            employee_name VARCHAR(255) NOT NULL, start_date DATE NOT NULL, end_date DATE NOT NULL,
            start_slot VARCHAR(12) NOT NULL, end_slot VARCHAR(12) NOT NULL,
            note TEXT NOT NULL, status VARCHAR(12) NOT NULL,
            created_at VARCHAR(40) NOT NULL, updated_at VARCHAR(40) NOT NULL {indexes})''')
        indexes = '' if store.sqlite else ', KEY leave_event_request (request_id,id)'
        # No ON DELETE CASCADE: operational history survives account deletion.
        store.execute(f'''CREATE TABLE IF NOT EXISTS leave_events (
            id {identity}, request_id BIGINT NOT NULL, actor_username VARCHAR(255) NOT NULL,
            actor_name VARCHAR(255) NOT NULL, action VARCHAR(20) NOT NULL,
            previous_status VARCHAR(12) NULL, new_status VARCHAR(12) NOT NULL,
            comment TEXT NOT NULL, created_at VARCHAR(40) NOT NULL {indexes})''')
        if store.sqlite:
            store.execute('CREATE INDEX IF NOT EXISTS leave_owner ON leave_requests(employee_username,status,id)')
            store.execute('CREATE INDEX IF NOT EXISTS leave_event_request ON leave_events(request_id,id)')
        else:
            align_leave_identity_collations(store)


def align_leave_identity_collations(store):
    """Match only leave identity columns to users.username; preserve legacy users and rows."""
    store.execute('''SELECT CHARACTER_SET_NAME AS charset, COLLATION_NAME AS collation
        FROM information_schema.COLUMNS WHERE TABLE_SCHEMA=DATABASE()
        AND TABLE_NAME='users' AND COLUMN_NAME='username' ''')
    target = store.one()
    if not target or not all(re.fullmatch(r'[A-Za-z0-9_]+', str(target.get(key) or '')) for key in ('charset', 'collation')):
        raise RuntimeError('Cannot determine the users.username character set/collation.')
    for table, column, nullable in [('leave_requests', 'employee_username', 'NULL'),
                                     ('leave_events', 'actor_username', 'NOT NULL')]:
        store.execute('''SELECT CHARACTER_SET_NAME AS charset, COLLATION_NAME AS collation
            FROM information_schema.COLUMNS WHERE TABLE_SCHEMA=DATABASE()
            AND TABLE_NAME=%s AND COLUMN_NAME=%s''', (table, column))
        current = store.one()
        if not current:
            raise RuntimeError('Leave identity column is missing.')
        if current != target:
            # Identifiers come from fixed module constants; metadata values are validated above.
            store.execute(f'ALTER TABLE {table} MODIFY COLUMN {column} VARCHAR(255) '
                          f'CHARACTER SET {target["charset"]} COLLATE {target["collation"]} {nullable}')


def rename_leave_owner(cursor, previous, current):
    """Follow account renames; immutable event actor snapshots remain unchanged."""
    cursor.execute('UPDATE leave_requests SET employee_username=%s WHERE employee_username=%s', (current, previous))


def detach_leave_owner(cursor, username):
    """Preserve history but prevent a recreated username inheriting old requests."""
    cursor.execute('UPDATE leave_requests SET employee_username=NULL WHERE employee_username=%s', (username,))
