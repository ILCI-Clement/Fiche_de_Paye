"""Transactional date synchronization; callers lock users before employee records."""
from datetime import date

from fastapi import HTTPException

try:
    from .personnel_rules import paris_today
    from .personnel_store import append_personnel_event, insert_period, load_employee_detail, touch_employee
except ImportError:
    from personnel_rules import paris_today
    from personnel_store import append_personnel_event, insert_period, load_employee_detail, touch_employee


def sync_current_contract(cursor, username):
    today = paris_today()
    cursor.execute('''SELECT start_date,end_date,is_cdi FROM personnel_contracts
        WHERE username=%s AND is_archived=0 AND start_date<=%s
        AND (end_date IS NULL OR end_date>=%s) ORDER BY start_date DESC,id DESC LIMIT 1''',
                   (username, today, today))
    current = cursor.fetchone()
    cursor.execute('UPDATE users SET contract_start_date=%s,contract_end_date=%s,is_cdi=%s WHERE username=%s',
                   (current['start_date'] if current else None, current['end_date'] if current else None,
                    bool(current and current['is_cdi']), username))


def detach_contract_links(cursor, employee_id):
    # Dates and their audit survive unlinking or reassignment to another account.
    cursor.execute('UPDATE personnel_register_periods SET account_contract_id=NULL WHERE employee_id=%s',
                   (employee_id,))


def sync_account_contracts(cursor, username, actor_username, *, match_legacy_period=True):
    """Account dates win on first association; stable IDs drive subsequent updates.

    Match old periods by start date, then only an unambiguous one-to-one remainder.
    Unmatched register history is retained, never guessed away or deleted.
    """
    cursor.execute('SELECT employee_id,registration_schema_version FROM personnel_register WHERE linked_username=%s FOR UPDATE', (username,))
    owner = cursor.fetchone()
    if not owner:
        return
    employee_id = owner['employee_id']
    match_legacy_period = match_legacy_period and owner['registration_schema_version'] < 2
    before = load_employee_detail(cursor, employee_id)
    cursor.execute('SELECT * FROM personnel_contracts WHERE username=%s ORDER BY start_date,id FOR UPDATE',
                   (username,))
    contracts = cursor.fetchall()
    cursor.execute('SELECT * FROM personnel_register_periods WHERE employee_id=%s ORDER BY period_id FOR UPDATE',
                   (employee_id,))
    periods = cursor.fetchall()
    assigned = {p['account_contract_id']: p for p in periods if p['account_contract_id'] is not None}
    remaining = [c for c in contracts if c['id'] not in assigned]
    free = [p for p in periods if p['account_contract_id'] is None]
    changed = False
    for contract in remaining:
        matches = [p for p in free if p['entry_date'] == contract['start_date']]
        if len(matches) == 1 and sum(c['start_date'] == contract['start_date'] for c in remaining) == 1:
            period = matches[0]
            assigned[contract['id']] = period
            free.remove(period)
    remaining = [c for c in contracts if c['id'] not in assigned]
    if match_legacy_period and len(remaining) == len(free) == 1:
        assigned[remaining[0]['id']] = free[0]
    for contract in contracts:
        period = assigned.get(contract['id'])
        if period and period.get('actual_exit_date') and contract['start_date'] > period['actual_exit_date']:
            raise HTTPException(409, 'Le début du contrat dépasserait la sortie effective. Vérifiez la période dans le registre.')
        if period is None:
            kind = 'CDI' if contract['is_cdi'] else 'stage' if contract['employee_type'] == 'stagiaire' else ''
            period_id = insert_period(cursor, employee_id, {
                'entry_date': contract['start_date'], 'exit_date': contract['end_date'],
                'contract_type': kind, 'other_contract': '',
            })
            cursor.execute('UPDATE personnel_register_periods SET account_contract_id=%s WHERE period_id=%s',
                           (contract['id'], period_id))
            changed = True
        elif (period['account_contract_id'] != contract['id'] or period['entry_date'] != contract['start_date']
              or period['exit_date'] != contract['end_date']):
            cursor.execute('''UPDATE personnel_register_periods SET account_contract_id=%s,entry_date=%s,exit_date=%s
                WHERE period_id=%s AND employee_id=%s''',
                           (contract['id'], contract['start_date'], contract['end_date'], period['period_id'], employee_id))
            changed = True
    if changed:
        touch_employee(cursor, employee_id)
        append_personnel_event(cursor, employee_id, actor_username, 'contract_dates_synced', before,
                               load_employee_detail(cursor, employee_id))


def sync_register_period(cursor, employee_id, period_id):
    """Push one explicit register edit to its account contract in the same transaction."""
    cursor.execute('SELECT linked_username FROM personnel_register WHERE employee_id=%s', (employee_id,))
    username = cursor.fetchone()['linked_username']
    if not username:
        return
    cursor.execute('SELECT * FROM personnel_register_periods WHERE employee_id=%s AND period_id=%s FOR UPDATE',
                   (employee_id, period_id))
    period = cursor.fetchone()
    contract_id = period['account_contract_id']
    if contract_id is not None:
        cursor.execute('SELECT * FROM personnel_contracts WHERE id=%s AND username=%s FOR UPDATE', (contract_id, username))
        contract = cursor.fetchone()
        if not contract:
            raise HTTPException(409, 'Le contrat associé a changé. Actualisez la fiche.')
        if contract['is_archived']:
            raise HTTPException(409, 'Restaurez le contrat dans Administration avant de modifier ses dates.')
        kind, cdi = contract['employee_type'], bool(contract['is_cdi'])
        if period['contract_type'] and (period['contract_type'] == 'CDI') != cdi:
            raise HTTPException(400, 'Le type CDI se modifie dans Administration pour le contrat associé.')
    else:
        cursor.execute('SELECT employee_type FROM users WHERE username=%s', (username,))
        account = cursor.fetchone()
        kind = account['employee_type'] or 'salarie'
        cdi = period['contract_type'] == 'CDI'
    start, end = period['entry_date'], period['exit_date']
    # Registering a person does not require inventing a fixed-term account contract.
    # This episode stays independent until its planned contract dates are complete.
    if contract_id is None and (start is None or (end is None and not cdi)):
        return
    if start is None or (end is None and not cdi):
        raise HTTPException(400, 'Pour synchroniser le compte, renseignez le début et la fin du contrat (sauf CDI).')
    if cdi and end is not None:
        raise HTTPException(400, 'Ce contrat est un CDI sans date de fin. Modifiez son type dans Administration.')
    if kind == 'stagiaire' and (cdi or end is None):
        raise HTTPException(400, 'Un stage doit avoir une date de fin.')
    if end is not None and end < start:
        raise HTTPException(400, 'La fin ne peut pas précéder le début.')
    cursor.execute('''SELECT id FROM personnel_contracts WHERE username=%s AND id<>%s AND is_archived=0
        AND start_date<=%s AND (end_date IS NULL OR end_date>=%s) FOR UPDATE''',
                   (username, contract_id or 0, end or date.max, start))
    if cursor.fetchone():
        raise HTTPException(409, 'Ces dates chevauchent un autre contrat du compte associé.')
    if contract_id is None:
        cursor.execute('''INSERT INTO personnel_contracts (username,employee_type,start_date,end_date,is_cdi)
            VALUES (%s,%s,%s,%s,%s)''', (username, kind, start, end, cdi))
        contract_id = cursor.lastrowid
        cursor.execute('UPDATE personnel_register_periods SET account_contract_id=%s WHERE period_id=%s',
                       (contract_id, period_id))
    else:
        cursor.execute('UPDATE personnel_contracts SET start_date=%s,end_date=%s WHERE id=%s AND username=%s',
                       (start, end, contract_id, username))
    sync_current_contract(cursor, username)
