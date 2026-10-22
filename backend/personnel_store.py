"""MySQL storage for employee identities, employment episodes and append-only audit."""
import json
import re
from datetime import date, datetime, timezone

try:
    from .personnel_rules import employment_status
except ImportError:
    from personnel_rules import employment_status


EMPLOYEE_FIELDS = (
    'full_name', 'nationality', 'sex', 'address', 'birth_date',
    'job_title', 'qualification', 'personnel_type', 'linked_username',
)
PERIOD_FIELDS = ('entry_date', 'exit_date', 'contract_type', 'other_contract')
EMPLOYEE_COLUMNS = ('employee_id', *EMPLOYEE_FIELDS, 'is_archived', 'version', 'created_at', 'updated_at',
                    'has_social_security_number')
PERIOD_COLUMNS = ('period_id', 'employee_id', *PERIOD_FIELDS, 'created_at', 'updated_at')
# Known dates take precedence; ties and entirely undated history use registration order.
PERIOD_ORDER = 'entry_date DESC, period_id DESC'


def selected_columns(columns, *, table='', prefix=''):
    """Read TIMESTAMP as its UTC epoch; caller/session time zones cannot shift audit dates."""
    expressions = []
    for column in columns:
        source = f'{table}.{column}' if table else column
        if column == 'has_social_security_number':
            encrypted = f'{table}.social_security_number_encrypted' if table else 'social_security_number_encrypted'
            # NULL is the only absence marker; malformed/empty envelopes must fail closed on reveal.
            source = f'({encrypted} IS NOT NULL)'
        elif column in {'created_at', 'updated_at'}:
            source = f'UNIX_TIMESTAMP({source})'
        expressions.append(f'{source} AS {prefix}{column}')
    return ','.join(expressions)


def utc_timestamp(value):
    return datetime.fromtimestamp(float(value), timezone.utc) if value is not None else None


def period_record(row):
    value = dict(row)
    for key in ('created_at', 'updated_at'):
        value[key] = utc_timestamp(value[key])
    return value


def ensure_personnel_schema(cursor):
    """Only add empty tables; never import existing users into official employee records."""
    # Match the legacy username's case/accent semantics for both the join and unique key.
    # These identifiers come from MySQL metadata, never from the request body.
    cursor.execute('''SELECT CHARACTER_SET_NAME AS charset_name,COLLATION_NAME AS collation_name
        FROM information_schema.columns WHERE table_schema=DATABASE()
        AND table_name='users' AND column_name='username' ''')
    metadata = cursor.fetchone()
    if not metadata or not all(re.fullmatch(r'[A-Za-z0-9_]+', metadata.get(key) or '')
                               for key in ('charset_name', 'collation_name')):
        raise RuntimeError('Cannot identify the account username collation for the personnel register.')
    account_text = f"CHARACTER SET {metadata['charset_name']} COLLATE {metadata['collation_name']}"
    cursor.execute(f'''CREATE TABLE IF NOT EXISTS personnel_register (
        employee_id BIGINT NOT NULL AUTO_INCREMENT PRIMARY KEY,
        full_name VARCHAR(200) NOT NULL,
        nationality VARCHAR(100) NOT NULL DEFAULT '',
        sex VARCHAR(1) NOT NULL DEFAULT '',
        address VARCHAR(1000) NOT NULL DEFAULT '',
        birth_date DATE NULL,
        job_title VARCHAR(200) NOT NULL DEFAULT '',
        qualification VARCHAR(200) NOT NULL DEFAULT '',
        personnel_type VARCHAR(20) NOT NULL DEFAULT '',
        linked_username VARCHAR(255) {account_text} NULL,
        social_security_number_encrypted TEXT NULL,
        is_archived TINYINT(1) NOT NULL DEFAULT 0,
        version INT NOT NULL DEFAULT 1,
        created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
        UNIQUE KEY personnel_register_account (linked_username),
        KEY personnel_register_filter (is_archived, personnel_type, employee_id)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4''')
    # SHOW COLUMNS also sees connection-scoped TEMPORARY tables used by isolated tests.
    cursor.execute("SHOW COLUMNS FROM personnel_register LIKE 'social_security_number_encrypted'")
    if not cursor.fetchone():
        cursor.execute('ALTER TABLE personnel_register ADD COLUMN social_security_number_encrypted TEXT NULL')
    cursor.execute('''CREATE TABLE IF NOT EXISTS personnel_register_periods (
        period_id BIGINT NOT NULL AUTO_INCREMENT PRIMARY KEY,
        employee_id BIGINT NOT NULL,
        entry_date DATE NULL,
        exit_date DATE NULL,
        contract_type VARCHAR(20) NOT NULL DEFAULT '',
        other_contract VARCHAR(200) NOT NULL DEFAULT '',
        created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
        KEY personnel_period_order (employee_id, entry_date, period_id),
        CONSTRAINT personnel_period_owner FOREIGN KEY (employee_id)
            REFERENCES personnel_register (employee_id) ON DELETE RESTRICT
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS personnel_register_events (
        event_id BIGINT NOT NULL AUTO_INCREMENT PRIMARY KEY,
        employee_id BIGINT NOT NULL,
        actor_username VARCHAR(255) NOT NULL,
        action VARCHAR(40) NOT NULL,
        before_snapshot LONGTEXT NULL,
        after_snapshot LONGTEXT NOT NULL,
        created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
        KEY personnel_event_order (employee_id, event_id),
        CONSTRAINT personnel_event_owner FOREIGN KEY (employee_id)
            REFERENCES personnel_register (employee_id) ON DELETE RESTRICT
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4''')


def employee_record(row, latest_period=None):
    value = {key: row[key] for key in EMPLOYEE_COLUMNS}
    value['registration_number'] = value['employee_id']
    value['is_archived'] = bool(value['is_archived'])
    value['has_social_security_number'] = bool(value['has_social_security_number'])
    for key in ('created_at', 'updated_at'):
        value[key] = utc_timestamp(value[key])
    value['latest_period'] = latest_period
    value['status'] = employment_status(latest_period)
    return value


def load_employee_detail(cursor, employee_id, *, lock=False):
    cursor.execute(
        f'SELECT {selected_columns(EMPLOYEE_COLUMNS)} FROM personnel_register WHERE employee_id=%s'
        + (' FOR UPDATE' if lock else ''), (employee_id,),
    )
    row = cursor.fetchone()
    if not row:
        return None
    cursor.execute(
        f'SELECT {selected_columns(PERIOD_COLUMNS)} FROM personnel_register_periods '
        f'WHERE employee_id=%s ORDER BY {PERIOD_ORDER}', (employee_id,),
    )
    periods = [period_record(row) for row in cursor.fetchall()]
    return {'employee': employee_record(row, periods[0] if periods else None), 'periods': periods}


def _json_value(value):
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    raise TypeError(f'Unsupported audit value: {type(value).__name__}')


def append_personnel_event(cursor, employee_id, actor_username, action, before, after):
    # Never log snapshots; sensitive address and date-of-birth stay inside Admin-only storage.
    cursor.execute('''INSERT INTO personnel_register_events
        (employee_id,actor_username,action,before_snapshot,after_snapshot)
        VALUES (%s,%s,%s,%s,%s)''', (
        employee_id, actor_username, action,
        None if before is None else json.dumps(safe_audit_snapshot(before), ensure_ascii=False, default=_json_value),
        json.dumps(safe_audit_snapshot(after), ensure_ascii=False, default=_json_value),
    ))


def safe_audit_snapshot(value):
    """Keep the presence flag, never a dedicated number or its encryption envelope."""
    if isinstance(value, dict):
        return {key: safe_audit_snapshot(item) for key, item in value.items()
                if key not in {'social_security_number', 'social_security_number_encrypted'}}
    if isinstance(value, (list, tuple)):
        return [safe_audit_snapshot(item) for item in value]
    return value


def sensitive_metadata(employee_id, has_number):
    """NIR events deliberately omit ordinary identity fields and all number material."""
    return {'employee': {'employee_id': employee_id, 'has_social_security_number': bool(has_number)},
            'periods': []}


def touch_employee(cursor, employee_id):
    cursor.execute('UPDATE personnel_register SET version=version+1 WHERE employee_id=%s', (employee_id,))


def delete_personnel_records(cursor, employee_id):
    """Delete this independent registry identity only, within its caller's transaction."""
    counts = {}
    for table, label in (('personnel_register_periods', 'period_count'),
                         ('personnel_register_events', 'history_count')):
        cursor.execute(f'SELECT COUNT(*) AS count FROM {table} WHERE employee_id=%s', (employee_id,))
        counts[label] = int(cursor.fetchone()['count'])
    cursor.execute('DELETE FROM personnel_register_events WHERE employee_id=%s', (employee_id,))
    cursor.execute('DELETE FROM personnel_register_periods WHERE employee_id=%s', (employee_id,))
    cursor.execute('DELETE FROM personnel_register WHERE employee_id=%s', (employee_id,))
    return {**counts, 'deleted_records': 1 + sum(counts.values()),
            'confirmed': True, 'previous_is_archived': True}


def insert_period(cursor, employee_id, values):
    cursor.execute(
        f"INSERT INTO personnel_register_periods (employee_id,{','.join(PERIOD_FIELDS)}) "
        'VALUES (%s,%s,%s,%s,%s)',
        (employee_id, *(values[key] for key in PERIOD_FIELDS)),
    )
    return cursor.lastrowid


def rename_personnel_account_link(cursor, old_username, new_username, actor_username):
    """Join the caller's rename transaction; lock all users before calling this helper.

    Historical audit actors and snapshots retain their original spelling. Only the current
    account link changes, with its own complete before/after audit and version increment.
    A legacy database without the register remains compatible.
    """
    cursor.execute('''SELECT 1 AS present FROM information_schema.tables
        WHERE table_schema=DATABASE() AND table_name='personnel_register' LIMIT 1''')
    if not cursor.fetchone():
        return
    cursor.execute('SELECT employee_id FROM personnel_register WHERE linked_username=%s FOR UPDATE', (old_username,))
    row = cursor.fetchone()
    if not row:
        return
    employee_id = row['employee_id']
    before = load_employee_detail(cursor, employee_id, lock=True)
    if before['employee']['linked_username'] == new_username:
        return
    cursor.execute('UPDATE personnel_register SET linked_username=%s,version=version+1 WHERE employee_id=%s',
                   (new_username, employee_id))
    after = load_employee_detail(cursor, employee_id)
    append_personnel_event(cursor, employee_id, actor_username, 'account_renamed', before, after)
