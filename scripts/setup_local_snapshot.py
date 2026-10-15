"""Initialize a new loopback test database from the private production snapshot once."""
import hashlib
import json
import secrets
import subprocess
from pathlib import Path

import pymysql

ROOT = Path(__file__).resolve().parents[1]
STATE = ROOT / '.local-test' / 'rebuild'
CONFIG = STATE / 'runtime.json'
DATABASE = 'presence_local_rebuild'
MYSQL = Path(r'C:\Program Files\MySQL\MySQL Server 8.4\bin\mysql.exe')


def setup():
    if CONFIG.exists():
        raise SystemExit('Local configuration already exists. Refusing to replace test data.')
    db = pymysql.connect(host='127.0.0.1', port=3307, user='root', password='', autocommit=True)
    try:
        with db.cursor() as cursor:
            cursor.execute('SHOW DATABASES LIKE %s', (DATABASE,))
            if cursor.fetchone():
                raise SystemExit('Test database already exists. Refusing to overwrite it.')
            cursor.execute(f'CREATE DATABASE {DATABASE} CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci')
        sql = (STATE / 'snapshot' / 'database.sql').read_bytes()
        # MariaDB's dump sandbox preamble is not understood by the MySQL client.
        sql = b'\n'.join(line for line in sql.splitlines() if not line.startswith(b'/*M!999999'))
        result = subprocess.run([str(MYSQL), '--no-defaults', '--host=127.0.0.1', '--port=3307',
            '--user=root', '--default-character-set=utf8mb4', DATABASE], input=sql, capture_output=True)
        if result.returncode:
            (STATE / 'import-errors.txt').write_bytes(result.stderr)
            raise SystemExit('Snapshot import failed. Private diagnostics saved locally; no production writes.')
        password = secrets.token_urlsafe(32)
        with db.cursor() as cursor:
            cursor.execute(f'UPDATE {DATABASE}.users SET remember_token_hash=NULL, remember_token_expires=NULL, '
                           'reset_token=NULL, reset_token_expires=%s', ('2000-01-01 00:00:00',))
            cursor.execute("ALTER USER 'root'@'localhost' IDENTIFIED BY %s", (password,))
        CONFIG.write_text(json.dumps({'db_password': password, 'api_token': secrets.token_urlsafe(48),
            'database': DATABASE, 'snapshot_sha256': hashlib.sha256(sql).hexdigest()}, indent=2), encoding='utf-8')
        print('Independent database imported. Local credentials stored in the ignored runtime file.')
    finally:
        db.close()


if __name__ == '__main__':
    setup()
