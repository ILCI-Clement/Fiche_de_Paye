"""Administrator-maintained, non-overlapping personnel contracts."""
from datetime import date
from calendar import monthrange
import json
from typing import Literal
from fastapi import Depends, HTTPException
from pydantic import BaseModel, ConfigDict, Field, StrictBool
try:
    from .personnel_contract_sync import sync_account_contracts, sync_current_contract
    from .destructive_audit import append_deletion_event
    from .role_policy import canonical_role_tags, checked_resource_actor, is_admin, is_super_admin
except ImportError:
    from personnel_contract_sync import sync_account_contracts, sync_current_contract
    from destructive_audit import append_deletion_event
    from role_policy import canonical_role_tags, checked_resource_actor, is_admin, is_super_admin


class ContractInput(BaseModel):
    employee_type: Literal['salarie', 'stagiaire', 'alternance']
    start_date: date
    end_date: date | None = None
    is_cdi: bool = False


class AttendanceContractInput(ContractInput):
    model_config = ConfigDict(extra='forbid')
    month: int = Field(ge=1, le=12)
    year: int = Field(ge=2000, le=2100)


class ContractArchiveInput(BaseModel):
    model_config = ConfigDict(extra='forbid')
    is_archived: StrictBool


class ContractDeleteInput(BaseModel):
    model_config = ConfigDict(extra='forbid', str_strip_whitespace=True)
    confirmation: str = Field(strict=True, min_length=1, max_length=32)


def contract_referenced_by_attendance(cursor, contract_id, username):
    """Inspect all workspaces; a workspace owner is not the employee on its sheets."""
    def contains_reference(value):
        if isinstance(value, dict):
            snapshot = value.get('contract_snapshot')
            snapshot = snapshot if isinstance(snapshot, dict) else {}
            identifiers = (value.get('contract_id'), snapshot.get('id'))
            if any(str(identifier) == str(contract_id) for identifier in identifiers if identifier is not None):
                owners = {str(owner).casefold() for owner in
                          (value.get('account_username'), snapshot.get('username')) if owner}
                # A legacy reference without an owner cannot safely be discarded.
                if not owners or str(username).casefold() in owners:
                    return True
            return any(contains_reference(item) for item in value.values())
        if isinstance(value, list):
            return any(contains_reference(item) for item in value)
        return False
    cursor.execute('SELECT form_content FROM Presence FOR UPDATE')
    for row in cursor.fetchall():
        raw = row['form_content']
        try:
            configuration = json.loads(raw) if isinstance(raw, (str, bytes, bytearray)) else raw
        except (ValueError, TypeError):
            raise HTTPException(409, 'Une archive de présence doit être vérifiée avant cette suppression.') from None
        if not isinstance(configuration, (dict, list)):
            raise HTTPException(409, 'Une archive de présence doit être vérifiée avant cette suppression.')
        if contains_reference(configuration):
            return True
    return False


def register_contract_history(app, connect, require_roles, get_current_user, access, *, load_user=None):
    def load_account(cursor, username):
        if load_user is not None:
            return load_user(cursor, username)
        cursor.execute('''SELECT username,role,role_tags,is_admin,is_active,session_version,login_deleted_at
            FROM users WHERE username=%s''', (username,))
        account = cursor.fetchone()
        if account:
            tags = account.get('role_tags')
            if isinstance(tags, str):
                try:
                    tags = json.loads(tags)
                except (ValueError, TypeError):
                    tags = []
            account['role_tags'] = canonical_role_tags(tags) or canonical_role_tags(
                [account.get('role') or ('Admin' if account.get('is_admin') else 'Employe')])
        return account

    def locked_admin(cursor, actor, *, super_only=False):
        current = checked_resource_actor(cursor, actor, load_account, lock=True)
        if not is_admin(current) or (super_only and not is_super_admin(current)):
            raise HTTPException(403, 'Cette opération est réservée au SuperAdmin.' if super_only
                                else 'Cette opération est réservée aux administrateurs.')
        return current

    def contract_target(cursor, username, contract_id):
        cursor.execute('SELECT username FROM users WHERE username=%s FOR UPDATE', (username,))
        target = cursor.fetchone()
        if not target:
            raise HTTPException(404, 'Personne introuvable.')
        canonical = target['username']
        cursor.execute('SELECT * FROM personnel_contracts WHERE id=%s AND username=%s FOR UPDATE',
                       (contract_id, canonical))
        contract = cursor.fetchone()
        if not contract:
            raise HTTPException(404, 'Contrat introuvable.')
        return canonical, contract

    @app.on_event('startup')
    def migrate_contract_history():
        db = connect()
        try:
            with db.cursor() as cur:
                cur.execute('''CREATE TABLE IF NOT EXISTS personnel_contracts (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    username VARCHAR(255) NOT NULL,
                    employee_type VARCHAR(20) NOT NULL,
                    start_date DATE NOT NULL,
                    end_date DATE NULL,
                    is_cdi BOOLEAN NOT NULL DEFAULT 0,
                    is_archived TINYINT(1) NOT NULL DEFAULT 0,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    KEY contracts_person_dates (username, start_date, end_date)
                )''')
                cur.execute("SHOW COLUMNS FROM personnel_contracts LIKE 'is_archived'")
                if not cur.fetchone():
                    cur.execute('ALTER TABLE personnel_contracts ADD COLUMN is_archived TINYINT(1) NOT NULL DEFAULT 0')
                cur.execute('''INSERT INTO personnel_contracts (username, employee_type, start_date, end_date, is_cdi)
                    SELECT u.username, COALESCE(u.employee_type, 'salarie'), u.contract_start_date,
                           IF(u.is_cdi, NULL, u.contract_end_date), u.is_cdi
                    FROM users u WHERE u.contract_start_date IS NOT NULL
                    AND (u.is_cdi OR u.contract_end_date >= u.contract_start_date)
                    AND NOT EXISTS (SELECT 1 FROM personnel_contracts c WHERE c.username=u.username)''')
                cur.execute('''CREATE TABLE IF NOT EXISTS contract_setup_events (
                    id BIGINT AUTO_INCREMENT PRIMARY KEY,
                    contract_id BIGINT NOT NULL, employee_username VARCHAR(255) NOT NULL,
                    actor_username VARCHAR(255) NOT NULL, fiche_month INT NOT NULL, fiche_year INT NOT NULL,
                    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP)''')
            db.commit()
        finally:
            db.close()

    @app.get('/users/{username}/contracts')
    def get_contracts(username: str, actor=Depends(get_current_user)):
        db = connect()
        try:
            with db.cursor() as cur:
                current = checked_resource_actor(cur, actor, load_account)
                access(cur, current, username, edit=False)
                cur.execute('SELECT * FROM personnel_contracts WHERE username=%s ORDER BY start_date,id', (username,))
                return {'contracts': [{**row, 'is_archived': bool(row.get('is_archived'))} for row in cur.fetchall()]}
        finally:
            db.close()

    def save(username, payload, contract_id, actor, *, attendance_period=None):
        if payload.is_cdi and payload.employee_type == 'stagiaire':
            raise HTTPException(400, 'Un stage doit avoir une date de fin.')
        end = None if payload.is_cdi else payload.end_date
        if not payload.is_cdi and end is None:
            raise HTTPException(400, 'La date de fin est obligatoire hors CDI.')
        if end and end < payload.start_date:
            raise HTTPException(400, 'La fin ne peut pas précéder le début.')
        if attendance_period:
            month, year = attendance_period
            first, last = date(year, month, 1), date(year, month, monthrange(year, month)[1])
            if payload.start_date > last or (end is not None and end < first):
                raise HTTPException(400, 'Le contrat doit couvrir le mois sélectionné.')
        db = connect()
        try:
            with db.cursor() as cur:
                if attendance_period:
                    current_actor = checked_resource_actor(cur, actor, load_account, lock=True)
                    target = load_account(cur, username)
                    if not target or not target.get('is_active', True) or target.get('login_deleted_at'):
                        raise HTTPException(404, 'Compte actif introuvable.')
                    if not is_admin(current_actor) and not (
                            'Responsable' in current_actor['role_tags']
                            and 'Employe' in target['role_tags'] and not is_admin(target)
                            and target.get('manager_username') == current_actor['username']
                            and target['username'] != current_actor['username']):
                        raise HTTPException(403, 'Seul un Admin ou le responsable direct peut compléter ce contrat.')
                else:
                    current_actor = locked_admin(cur, actor)
                # Serialize edits per person to prevent two concurrent overlapping inserts.
                cur.execute('SELECT username FROM users WHERE username=%s FOR UPDATE', (username,))
                if not cur.fetchone():
                    raise HTTPException(404, 'Personne introuvable.')
                if attendance_period:
                    cur.execute('''SELECT id FROM personnel_contracts WHERE username=%s AND is_archived=0
                        AND start_date<=%s AND (end_date IS NULL OR end_date>=%s)''', (username, last, first))
                    if cur.fetchone():
                        raise HTTPException(409, 'Un contrat couvre déjà ce mois. Actualisez puis ouvrez la fiche.')
                archived = False
                if contract_id is not None:
                    cur.execute('SELECT id,is_archived FROM personnel_contracts WHERE id=%s AND username=%s FOR UPDATE', (contract_id, username))
                    existing = cur.fetchone()
                    if not existing:
                        raise HTTPException(404, 'Contrat introuvable.')
                    archived = bool(existing['is_archived'])
                    if archived:
                        raise HTTPException(409, 'Restaurez le contrat avant de modifier ses dates.')
                cur.execute('''SELECT id FROM personnel_contracts WHERE username=%s AND id<>%s
                    AND is_archived=0 AND start_date <= %s AND (end_date IS NULL OR end_date >= %s)''',
                    (username, contract_id or 0, end or date.max, payload.start_date))
                if cur.fetchone() and not archived:
                    raise HTTPException(409, 'Les dates chevauchent un autre contrat. Terminez le précédent avant de créer le nouveau.')
                values = (payload.employee_type, payload.start_date, end, payload.is_cdi)
                if contract_id is None:
                    cur.execute('INSERT INTO personnel_contracts (employee_type,start_date,end_date,is_cdi,username) VALUES (%s,%s,%s,%s,%s)', values+(username,))
                    contract_id = cur.lastrowid
                else:
                    cur.execute('UPDATE personnel_contracts SET employee_type=%s,start_date=%s,end_date=%s,is_cdi=%s WHERE id=%s AND username=%s', values+(contract_id,username))
                # Keep existing dashboard fields consistent with the contract active today.
                sync_current_contract(cur, username)
                sync_account_contracts(cur, username, current_actor['username'])
                if attendance_period:
                    cur.execute('''INSERT INTO contract_setup_events
                        (contract_id,employee_username,actor_username,fiche_month,fiche_year)
                        VALUES (%s,%s,%s,%s,%s)''', (contract_id, username, current_actor['username'], month, year))
            db.commit()
            return {'status':'success','contract_id':contract_id}
        except Exception:
            db.rollback()
            raise
        finally:
            db.close()

    @app.post('/users/{username}/contracts')
    def create_contract(username: str, payload: ContractInput, actor=Depends(require_roles('Admin'))):
        return save(username,payload,None,actor)

    @app.post('/attendance/contracts/{username}')
    def complete_attendance_contract(username: str, payload: AttendanceContractInput, actor=Depends(get_current_user)):
        # Creation only: existing contract modification/archive/delete stays Admin-only.
        return save(username, payload, None, actor, attendance_period=(payload.month, payload.year))

    @app.patch('/users/{username}/contracts/{contract_id}')
    def update_contract(username: str, contract_id: int, payload: ContractInput, actor=Depends(require_roles('Admin'))):
        return save(username,payload,contract_id,actor)

    @app.patch('/users/{username}/contracts/{contract_id}/archive')
    def archive_contract(username: str, contract_id: int, payload: ContractArchiveInput,
                         actor=Depends(get_current_user)):
        db = connect()
        try:
            with db.cursor() as cur:
                locked_admin(cur, actor)
                username, contract = contract_target(cur, username, contract_id)
                if not payload.is_archived and contract['is_archived']:
                    cur.execute('''SELECT id FROM personnel_contracts WHERE username=%s AND id<>%s AND is_archived=0
                        AND start_date<=%s AND (end_date IS NULL OR end_date>=%s)''',
                        (username, contract_id, contract['end_date'] or date.max, contract['start_date']))
                    if cur.fetchone():
                        raise HTTPException(409, 'Un contrat actif couvre déjà cette période.')
                cur.execute('UPDATE personnel_contracts SET is_archived=%s WHERE id=%s AND username=%s',
                            (payload.is_archived, contract_id, username))
                sync_current_contract(cur, username)
            db.commit()
            return {'status': 'success', 'contract_id': contract_id, 'is_archived': payload.is_archived}
        except Exception:
            db.rollback()
            raise
        finally:
            db.close()

    @app.delete('/users/{username}/contracts/{contract_id}')
    def delete_contract(username: str, contract_id: int, payload: ContractDeleteInput | None = None,
                        actor=Depends(get_current_user)):
        db = connect()
        try:
            with db.cursor() as cur:
                current = locked_admin(cur, actor, super_only=True)
                username, contract = contract_target(cur, username, contract_id)
                if not contract['is_archived']:
                    raise HTTPException(409, 'Archivez ce contrat avant sa suppression définitive.')
                if payload is None or payload.confirmation != str(contract_id):
                    raise HTTPException(400, 'Confirmez le numéro exact du contrat.')
                if contract_referenced_by_attendance(cur, contract_id, username):
                    raise HTTPException(409, 'Ce contrat est lié à des fiches de présence conservées ; sa suppression est refusée.')
                cur.execute('SELECT period_id FROM personnel_register_periods WHERE account_contract_id=%s LIMIT 1', (contract_id,))
                if cur.fetchone():
                    raise HTTPException(409, 'Ce contrat est lié au registre du personnel ; conservez-le archivé.')
                cur.execute('DELETE FROM personnel_contracts WHERE id=%s AND username=%s', (contract_id, username))
                sync_current_contract(cur, username)
                append_deletion_event(cur, current, 'contract', contract_id,
                                      {'deleted_records': 1, 'contract_count': 1,
                                       'confirmed': True, 'previous_is_archived': True})
            db.commit()
            return {'status': 'success', 'contract_id': contract_id}
        except Exception:
            db.rollback()
            raise
        finally:
            db.close()
