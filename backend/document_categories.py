"""Admin-managed document types; stable codes preserve historical receipt links."""
import hashlib
import json
import secrets
import unicodedata

import pymysql
from fastapi import Depends, HTTPException, Response
from pydantic import BaseModel, ConfigDict, Field, StrictBool, field_validator

from .role_policy import checked_resource_actor, is_admin

BUILTIN_TYPES = {'transport': 'Transport', 'autre': 'Autre'}


def normalized_label(value):
    value = ' '.join(unicodedata.normalize('NFKC', value).split())
    if not value or len(value) > 80 or any(unicodedata.category(char).startswith('C') for char in value):
        raise ValueError('Le libellé doit contenir entre 1 et 80 caractères lisibles.')
    if label_key(value) in {label_key('Tous'), label_key('Non renseigné')}:
        raise ValueError('Ce libellé est réservé aux filtres.')
    return value


def label_key(label):
    value = ''.join(char for char in unicodedata.normalize('NFKD', label).casefold()
                    if not unicodedata.combining(char))
    return hashlib.sha256(' '.join(value.split()).encode()).hexdigest()


class CategoryCreate(BaseModel):
    model_config = ConfigDict(extra='forbid')
    label: str = Field(min_length=1, max_length=160)

    @field_validator('label')
    @classmethod
    def label_value(cls, value):
        return normalized_label(value) if value is not None else value


class CategoryUpdate(CategoryCreate):
    label: str | None = Field(default=None, min_length=1, max_length=160)
    is_active: StrictBool | None = None
    expected_version: int = Field(strict=True, ge=1, le=2_147_483_646)


def ensure_category_schema(connect):
    db = connect()
    locked = False
    try:
        with db.cursor() as cursor:
            cursor.execute("SELECT GET_LOCK(CONCAT(DATABASE(), ':document_categories'), 60) AS acquired")
            locked = bool(cursor.fetchone()['acquired'])
            if not locked:
                raise RuntimeError('Document category migration lock unavailable')
            cursor.execute('''CREATE TABLE IF NOT EXISTS document_categories (
                code VARCHAR(20) CHARACTER SET ascii COLLATE ascii_bin NOT NULL PRIMARY KEY,
                label VARCHAR(80) NOT NULL,
                label_key CHAR(64) CHARACTER SET ascii COLLATE ascii_bin NOT NULL UNIQUE,
                is_active TINYINT(1) NOT NULL DEFAULT 1,
                is_system TINYINT(1) NOT NULL DEFAULT 0,
                version INT NOT NULL DEFAULT 1
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4''')
            cursor.execute('''CREATE TABLE IF NOT EXISTS document_category_events (
                id BIGINT NOT NULL AUTO_INCREMENT PRIMARY KEY,
                category_code VARCHAR(20) NOT NULL, receipt_id INT NULL,
                actor_username VARCHAR(255) NOT NULL, action VARCHAR(32) NOT NULL,
                before_json TEXT NULL, after_json TEXT NOT NULL,
                created_at DATETIME(6) NOT NULL DEFAULT CURRENT_TIMESTAMP(6),
                KEY document_category_audit (category_code,id),
                KEY document_reclassification (receipt_id,id)
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4''')
            for code, label in BUILTIN_TYPES.items():
                cursor.execute('''INSERT IGNORE INTO document_categories
                    (code,label,label_key,is_system) VALUES (%s,%s,%s,1)''', (code, label, label_key(label)))
        db.commit()
    except Exception:
        db.rollback()
        raise
    finally:
        if locked:
            with db.cursor() as cursor:
                cursor.execute("SELECT RELEASE_LOCK(CONCAT(DATABASE(), ':document_categories'))")
        db.close()


def category_record(row):
    return {'code': row['code'], 'label': row['label'], 'is_active': bool(row['is_active']),
            'is_system': bool(row['is_system']), 'version': int(row['version'])}


def category_event(cursor, actor, code, action, before, after, receipt_id=None):
    cursor.execute('''INSERT INTO document_category_events
        (category_code,receipt_id,actor_username,action,before_json,after_json) VALUES (%s,%s,%s,%s,%s,%s)''',
        (code, receipt_id, actor['username'], action,
         json.dumps(before, ensure_ascii=False) if before is not None else None, json.dumps(after, ensure_ascii=False)))


def resolve_category(cursor, code, other_type):
    cursor.execute('SELECT * FROM document_categories WHERE code=%s FOR UPDATE', (code,))
    row = cursor.fetchone()
    if not row:
        raise HTTPException(400, 'Choisissez un type de document existant.')
    if not row['is_active']:
        raise HTTPException(409, 'Ce type de document est désactivé. Actualisez et choisissez un autre type.')
    other = str(other_type or '').strip() if code == 'autre' else None
    if code == 'autre' and (not other or len(other) > 120):
        raise HTTPException(400, 'Précisez le type de document (1 à 120 caractères).')
    return row['label'], other


def receipt_category_label(row, labels=None):
    code = row.get('document_type')
    label = row.get('category_label') or (labels or {}).get(code) or BUILTIN_TYPES.get(code)
    if code == 'autre':
        return f"Autre · {row.get('other_type') or 'Non précisé'}"
    return label or 'Non renseigné'


def load_category_labels(cursor, rows):
    codes = sorted({r['document_type'] for r in rows if r.get('document_type') and r['document_type'] not in BUILTIN_TYPES})
    if not codes:
        return {}
    placeholders = ','.join(['%s'] * len(codes))
    cursor.execute(f'SELECT code,label FROM document_categories WHERE code IN ({placeholders})', tuple(codes))
    return {row['code']: row['label'] for row in cursor.fetchall()}


def register_category_routes(app, *, get_db_connection, get_current_user, load_user):
    app.router.add_event_handler('startup', lambda: ensure_category_schema(get_db_connection))

    @app.get('/document-categories')
    def categories(response: Response, actor=Depends(get_current_user)):
        response.headers['Cache-Control'] = 'no-store'
        db = get_db_connection()
        try:
            with db.cursor() as cursor:
                current = checked_resource_actor(cursor, actor, load_user)
                cursor.execute('SELECT * FROM document_categories ORDER BY is_system DESC,label,code')
                return {'categories': [category_record(row) for row in cursor.fetchall()], 'can_manage': is_admin(current)}
        finally:
            db.close()

    def admin(cursor, actor):
        current = checked_resource_actor(cursor, actor, load_user, lock=True)
        if not is_admin(current):
            raise HTTPException(403, 'La gestion des types est réservée aux administrateurs.')
        return current

    @app.post('/document-categories', status_code=201)
    def create(payload: CategoryCreate, actor=Depends(get_current_user)):
        db = get_db_connection()
        try:
            with db.cursor() as cursor:
                current = admin(cursor, actor)
                code = 'c_' + secrets.token_hex(9)  # Fits the existing VARCHAR(20) receipt field.
                cursor.execute('''INSERT INTO document_categories (code,label,label_key) VALUES (%s,%s,%s)''',
                               (code, payload.label, label_key(payload.label)))
                result = {'code': code, 'label': payload.label, 'is_active': True, 'is_system': False, 'version': 1}
                category_event(cursor, current, code, 'created', None, result)
            db.commit()
            return result
        except pymysql.err.IntegrityError as error:
            db.rollback()
            if error.args[0] == 1062:
                raise HTTPException(409, 'Ce type existe déjà, éventuellement dans les types désactivés.') from None
            raise
        except Exception:
            db.rollback()
            raise
        finally:
            db.close()

    @app.patch('/document-categories/{code}')
    def update(code: str, payload: CategoryUpdate, actor=Depends(get_current_user)):
        db = get_db_connection()
        try:
            with db.cursor() as cursor:
                current = admin(cursor, actor)
                cursor.execute('SELECT * FROM document_categories WHERE code=%s FOR UPDATE', (code,))
                row = cursor.fetchone()
                if not row:
                    raise HTTPException(404, 'Type introuvable.')
                before = category_record(row)
                if row['version'] != payload.expected_version:
                    raise HTTPException(409, 'Ce type a été modifié. Actualisez avant de réessayer.')
                if row['is_system']:
                    raise HTTPException(409, 'Transport et Autre sont conservés pour les documents existants.')
                if payload.label is None and payload.is_active is None:
                    raise HTTPException(400, 'Indiquez une modification.')
                label = payload.label if payload.label is not None else row['label']
                active = payload.is_active if payload.is_active is not None else bool(row['is_active'])
                result = {**before, 'label': label, 'is_active': active}
                if result != before:
                    result['version'] += 1
                    cursor.execute('''UPDATE document_categories SET label=%s,label_key=%s,is_active=%s,version=version+1
                        WHERE code=%s''', (label, label_key(label), active, code))
                    category_event(cursor, current, code, 'updated', before, result)
            db.commit()
            return result
        except pymysql.err.IntegrityError as error:
            db.rollback()
            if error.args[0] == 1062:
                raise HTTPException(409, 'Un autre type porte déjà ce libellé.') from None
            raise
        except Exception:
            db.rollback()
            raise
        finally:
            db.close()
