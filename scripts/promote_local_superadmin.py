"""Explicit, one-time role setup for the independent loopback test database only.

Never imported by an API entry point. Does not create accounts, set passwords or
print credentials. A private role-only backup is written before the transaction.
"""
import argparse
from datetime import datetime, timezone
import json
import os
from pathlib import Path
import sys
import pymysql

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from backend.role_policy import canonical_role_tags, is_super_admin, record_role_change


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--username', required=True)
    parser.add_argument('--expected-email', required=True)
    parser.add_argument('--confirm-local-only', action='store_true', required=True)
    args = parser.parse_args()
    state = ROOT / '.local-test' / 'rebuild'
    runtime = json.loads((state / 'runtime.json').read_text(encoding='utf-8-sig'))
    if runtime.get('database') != 'presence_local_rebuild':
        raise RuntimeError('Refusing a database other than the independent local test database.')
    db = pymysql.connect(host='127.0.0.1', port=3307, user='root', password=runtime['db_password'],
                         database='presence_local_rebuild', autocommit=False,
                         cursorclass=pymysql.cursors.DictCursor)
    try:
        with db.cursor() as cursor:
            cursor.execute('SELECT username FROM users ORDER BY username FOR UPDATE')
            cursor.fetchall()
            cursor.execute('''SELECT username,email,role,role_tags,is_admin,is_active,
                session_version,login_deleted_at FROM users WHERE username=%s''', (args.username,))
            target = cursor.fetchone()
            if (not target or target['username'] != args.username
                    or str(target['email']).strip().casefold() != args.expected_email.strip().casefold()
                    or not target['is_active'] or target['login_deleted_at']):
                raise RuntimeError('The exact active local test account could not be verified.')
            if is_super_admin(target):
                print('Local test account already has Super Admin; no change made.')
                return
            before = json.loads(target['role_tags']) if target['role_tags'] else [target['role']]
            after = canonical_role_tags([*before, 'SuperAdmin'])
            backup_dir = state / 'role-backups'
            backup_dir.mkdir(mode=0o700, parents=True, exist_ok=True)
            backup = {key: target[key] for key in ('username', 'role', 'role_tags', 'is_admin', 'session_version')}
            backup_file = backup_dir / (datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%S%fZ') + '.json')
            fd = os.open(backup_file, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
            with os.fdopen(fd, 'w', encoding='utf-8') as stream:
                json.dump(backup, stream)
            cursor.execute('''UPDATE users SET role='SuperAdmin',role_tags=%s,is_admin=1,
                session_version=session_version+1,remember_token_hash=NULL,remember_token_expires=NULL,
                reset_token=NULL,reset_token_expires=CURRENT_TIMESTAMP WHERE username=%s''',
                (json.dumps(after), target['username']))
            record_role_change(cursor, 'local-bootstrap', target['username'], before, after,
                               action='local_super_admin_granted')
        db.commit()
        print('Super Admin assigned to the verified local test account. Sign in again; production was not contacted.')
    except Exception:
        db.rollback()
        raise
    finally:
        db.close()


if __name__ == '__main__':
    main()
