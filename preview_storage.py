"""Local-only preview fixtures and persistence, independent of the Presence API."""

import json
import sqlite3
from contextlib import closing
from pathlib import Path

from app_core import ensure_workspace

PREVIEW_DATABASE = Path(__file__).resolve().parent / '.local-test' / 'preview.sqlite3'


DEMO_IDENTITIES = {
    ('employee', 'Camille', 'Exemple'),
    ('intern', 'Alex', 'Démo'),
}


def empty_workspace():
    """Return an empty local workspace without creating demonstration people."""
    return ensure_workspace({})


def remove_legacy_demo_records(workspace):
    """Remove the preview records shipped before local testing became user-driven."""
    workspace = ensure_workspace(workspace)

    def is_demo_record(record):
        details = record.get('person_snapshot', record)
        identity = (
            record.get('person_type', record.get('type')),
            record.get('first_name'),
            record.get('last_name'),
        )
        return identity in DEMO_IDENTITIES and details.get('supervisor_email') == 'responsable@example.invalid'

    demo_person_ids = {person.get('id') for person in workspace['people'] if is_demo_record(person)}
    workspace['people'] = [person for person in workspace['people'] if person.get('id') not in demo_person_ids]
    workspace['forms'] = [
        form for form in workspace['forms']
        if form.get('person_id') not in demo_person_ids and not is_demo_record(form)
    ]
    workspace['trash'] = [
        form for form in workspace['trash']
        if form.get('person_id') not in demo_person_ids and not is_demo_record(form)
    ]
    return workspace


def load_preview_workspace(database=PREVIEW_DATABASE):
    """Read only the local preview store; do not create a file until explicit save."""
    database = Path(database)
    if not database.exists():
        return empty_workspace()
    with closing(sqlite3.connect(database)) as connection:
        row = connection.execute('SELECT payload FROM workspace WHERE id = 1').fetchone()
    if not row:
        return empty_workspace()
    workspace = ensure_workspace(json.loads(row[0]))
    serialized_before_cleanup = json.dumps(workspace, ensure_ascii=False, sort_keys=True, default=str)
    cleaned_workspace = remove_legacy_demo_records(workspace)
    if json.dumps(cleaned_workspace, ensure_ascii=False, sort_keys=True, default=str) != serialized_before_cleanup:
        save_preview_workspace(cleaned_workspace, database)
    return cleaned_workspace


def save_preview_workspace(workspace, database=PREVIEW_DATABASE):
    """Commit a user-requested save locally using a transaction."""
    database = Path(database)
    database.parent.mkdir(parents=True, exist_ok=True)
    payload = json.dumps(workspace, ensure_ascii=False, default=str)
    with closing(sqlite3.connect(database)) as connection, connection:
        connection.execute('CREATE TABLE IF NOT EXISTS workspace (id INTEGER PRIMARY KEY, payload TEXT NOT NULL)')
        connection.execute('INSERT INTO workspace (id, payload) VALUES (1, ?) ON CONFLICT(id) DO UPDATE SET payload=excluded.payload', (payload,))
