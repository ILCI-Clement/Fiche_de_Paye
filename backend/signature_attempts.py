"""Durable local submission receipts: replay never invokes the signing provider twice."""
import json
import sqlite3
from contextlib import closing
from pathlib import Path
from fastapi import HTTPException


class SignatureAttempts:
    def __init__(self, path):
        self.path = Path(path)

    def connect(self):
        self.path.parent.mkdir(parents=True, exist_ok=True)
        db = sqlite3.connect(self.path, timeout=15)
        db.execute('''CREATE TABLE IF NOT EXISTS attempts (
            owner TEXT NOT NULL, request_id TEXT NOT NULL, fingerprint TEXT NOT NULL,
            status TEXT NOT NULL, result TEXT, created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            PRIMARY KEY(owner, request_id))''')
        return db

    def begin(self, owner, request_id, fingerprint):
        with closing(self.connect()) as db, db:
            try:
                db.execute('INSERT INTO attempts(owner,request_id,fingerprint,status) VALUES(?,?,?,?)',
                           (owner, request_id, fingerprint, 'pending'))
                return None
            except sqlite3.IntegrityError:
                row = db.execute('SELECT fingerprint,status,result FROM attempts WHERE owner=? AND request_id=?',
                                 (owner, request_id)).fetchone()
                if row[0] != fingerprint:
                    raise HTTPException(409, 'La fiche a changé. Actualisez la préparation avant tout nouvel envoi.')
                if row[1] != 'accepted':
                    raise HTTPException(409, 'Cette tentative a déjà été lancée. Vérifiez son résultat dans ClawShow avant de recommencer.')
                return json.loads(row[2])

    def finish(self, owner, request_id, status, result=None):
        with closing(self.connect()) as db, db:
            db.execute('UPDATE attempts SET status=?,result=? WHERE owner=? AND request_id=?',
                       (status, json.dumps(result, ensure_ascii=False) if result else None, owner, request_id))

    def profile_receipts(self, username, year):
        """Only linked acceptance metadata; never expose signing URLs or raw responses."""
        if not self.path.exists():
            return [], True
        try:
            with closing(sqlite3.connect(self.path.resolve().as_uri() + '?mode=ro', uri=True, timeout=15)) as db:
                rows = db.execute("SELECT result,created_at FROM attempts WHERE status='accepted' AND result IS NOT NULL").fetchall()
            receipts = []
            for raw, created in rows:
                value = json.loads(raw)
                subject = value.get('profile_subject') or {}
                month = subject.get('month')
                if subject.get('username') == username and subject.get('year') == year and isinstance(month, int) and 1 <= month <= 12:
                    if value.get('document_id'):
                        receipts.append({'month': month, 'document_id': str(value['document_id']), 'created_at': created,
                                         'mode': value.get('mode', 'unknown'), 'status': 'sent'})
            return receipts, True
        except (sqlite3.Error, ValueError, TypeError, AttributeError):
            return [], False
