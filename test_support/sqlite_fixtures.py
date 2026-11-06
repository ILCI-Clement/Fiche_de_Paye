"""Fast disposable SQLite fixtures; production durability settings are untouched.

These databases exercise SQL, constraints and transaction rollback, not recovery
after machine/power failure. Journals remain transactional but need no disk fsync.
Real MySQL tests retain their normal connection and durability settings.
"""
from pathlib import Path
import sqlite3
import tempfile


def connect_fixture(database, **options):
    memory = str(database) == ':memory:'
    if not memory and not Path(database).resolve().is_relative_to(Path(tempfile.gettempdir()).resolve()):
        raise ValueError('Disposable SQLite fixtures must stay inside the OS temporary directory.')
    db = sqlite3.connect(database, **options)
    if not memory:
        db.execute('PRAGMA journal_mode=MEMORY')
        db.execute('PRAGMA synchronous=OFF')
    return db
