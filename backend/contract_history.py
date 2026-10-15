"""Administrator-maintained, non-overlapping personnel contracts."""
from datetime import date
from typing import Literal
from fastapi import Depends, HTTPException
from pydantic import BaseModel


class ContractInput(BaseModel):
    employee_type: Literal['salarie', 'stagiaire', 'alternance']
    start_date: date
    end_date: date | None = None
    is_cdi: bool = False


def register_contract_history(app, connect, require_roles, get_current_user, access):
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
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    KEY contracts_person_dates (username, start_date, end_date)
                )''')
                cur.execute('''INSERT INTO personnel_contracts (username, employee_type, start_date, end_date, is_cdi)
                    SELECT u.username, COALESCE(u.employee_type, 'salarie'), u.contract_start_date,
                           IF(u.is_cdi, NULL, u.contract_end_date), u.is_cdi
                    FROM users u WHERE u.contract_start_date IS NOT NULL
                    AND (u.is_cdi OR u.contract_end_date >= u.contract_start_date)
                    AND NOT EXISTS (SELECT 1 FROM personnel_contracts c WHERE c.username=u.username)''')
            db.commit()
        finally:
            db.close()

    @app.get('/users/{username}/contracts')
    def get_contracts(username: str, actor=Depends(get_current_user)):
        db = connect()
        try:
            with db.cursor() as cur:
                access(cur, actor, username, edit=False)
                cur.execute('SELECT * FROM personnel_contracts WHERE username=%s ORDER BY start_date,id', (username,))
                return {'contracts': cur.fetchall()}
        finally:
            db.close()

    def save(username, payload, contract_id, actor):
        if payload.is_cdi and payload.employee_type == 'stagiaire':
            raise HTTPException(400, 'Un stage doit avoir une date de fin.')
        end = None if payload.is_cdi else payload.end_date
        if not payload.is_cdi and end is None:
            raise HTTPException(400, 'La date de fin est obligatoire hors CDI.')
        if end and end < payload.start_date:
            raise HTTPException(400, 'La fin ne peut pas précéder le début.')
        db = connect()
        try:
            with db.cursor() as cur:
                # Serialize edits per person to prevent two concurrent overlapping inserts.
                cur.execute('SELECT username FROM users WHERE username=%s FOR UPDATE', (username,))
                if not cur.fetchone():
                    raise HTTPException(404, 'Personne introuvable.')
                if contract_id is not None:
                    cur.execute('SELECT id FROM personnel_contracts WHERE id=%s AND username=%s', (contract_id, username))
                    if not cur.fetchone():
                        raise HTTPException(404, 'Contrat introuvable.')
                cur.execute('''SELECT id FROM personnel_contracts WHERE username=%s AND id<>%s
                    AND start_date <= %s AND (end_date IS NULL OR end_date >= %s)''',
                    (username, contract_id or 0, end or date.max, payload.start_date))
                if cur.fetchone():
                    raise HTTPException(409, 'Les dates chevauchent un autre contrat. Terminez le précédent avant de créer le nouveau.')
                values = (payload.employee_type, payload.start_date, end, payload.is_cdi)
                if contract_id is None:
                    cur.execute('INSERT INTO personnel_contracts (employee_type,start_date,end_date,is_cdi,username) VALUES (%s,%s,%s,%s,%s)', values+(username,))
                    contract_id = cur.lastrowid
                else:
                    cur.execute('UPDATE personnel_contracts SET employee_type=%s,start_date=%s,end_date=%s,is_cdi=%s WHERE id=%s AND username=%s', values+(contract_id,username))
                # Keep existing dashboard fields consistent with the contract active today.
                cur.execute('SELECT * FROM personnel_contracts WHERE username=%s AND start_date<=%s AND (end_date IS NULL OR end_date>=%s) ORDER BY start_date DESC LIMIT 1', (username,date.today(),date.today()))
                current = cur.fetchone()
                cur.execute('UPDATE users SET contract_start_date=%s,contract_end_date=%s,is_cdi=%s WHERE username=%s',
                    (current['start_date'] if current else None,current['end_date'] if current else None,current['is_cdi'] if current else False,username))
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

    @app.patch('/users/{username}/contracts/{contract_id}')
    def update_contract(username: str, contract_id: int, payload: ContractInput, actor=Depends(require_roles('Admin'))):
        return save(username,payload,contract_id,actor)
