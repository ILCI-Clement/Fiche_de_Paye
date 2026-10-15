"""Loopback-only leave preview. Never connect this identity selector to production."""
import sqlite3
import secrets
from pathlib import Path

from fastapi import FastAPI, Header, HTTPException
from pydantic import BaseModel

from .leave_requests import register_leave_routes

DATABASE = Path(__file__).resolve().parents[1] / '.local-test' / 'leave-preview.sqlite3'
ACCOUNTS = [
    {'username': 'alice.demo', 'name': 'Alice Martin', 'role': 'Employe', 'manager': 'robin.demo'},
    {'username': 'fabrice.demo', 'name': 'Fabrice Lam', 'role': 'Employe', 'manager': 'robin.demo'},
    {'username': 'robin.demo', 'name': 'Robin Hardy', 'role': 'Responsable', 'manager': 'william.demo'},
    {'username': 'william.demo', 'name': 'William Tang', 'role': 'Admin', 'manager': None},
    {'username': 'sans-responsable.demo', 'name': 'Compte sans responsable', 'role': 'Employe', 'manager': None},
]
TOKENS = {}
app = FastAPI(title='Presence leave — isolated local preview')


def connect():
    DATABASE.parent.mkdir(parents=True, exist_ok=True)
    db = sqlite3.connect(DATABASE, timeout=15)
    db.row_factory = sqlite3.Row
    return db


@app.on_event('startup')
def prepare_preview():
    db = connect()
    try:
        db.execute('''CREATE TABLE IF NOT EXISTS users (
            username TEXT PRIMARY KEY, email TEXT NOT NULL,
            manager_username TEXT, role TEXT NOT NULL)''')
        for account in ACCOUNTS:
            db.execute('INSERT OR IGNORE INTO users VALUES (?,?,?,?)',
                       (account['username'], account['name'], account['manager'], account['role']))
        db.commit()
    finally:
        db.close()


def actor(authorization: str | None = Header(default=None)):
    username = TOKENS.get(str(authorization or '').removeprefix('Bearer '))
    if not username:
        raise HTTPException(401, 'Connectez-vous avec un compte de démonstration.')
    db = connect()
    try:
        row = db.execute('SELECT * FROM users WHERE username=?', (username,)).fetchone()
        if not row:
            raise HTTPException(401, 'Session invalide.')
        return {**dict(row), 'role_tags': [row['role']]}
    finally:
        db.close()


class PreviewLogin(BaseModel):
    username: str


@app.get('/preview/accounts')
def accounts():
    return ACCOUNTS


@app.post('/preview/login')
def preview_login(payload: PreviewLogin):
    if not any(account['username'] == payload.username for account in ACCOUNTS):
        raise HTTPException(401, 'Compte de démonstration inconnu.')
    token = secrets.token_urlsafe(32)
    TOKENS[token] = payload.username
    return {'auth_token': token}


register_leave_routes(app, get_db_connection=connect, get_current_user=actor)
