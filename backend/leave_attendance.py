"""Atomic leave-to-attendance synchronization; PDFs and signed source data are immutable.

All writes use the caller's leave transaction. The cell ledger stores only calendar
state, not a whole person/workspace snapshot, so cancellation cannot undo unrelated edits.
"""
import copy
import hashlib
import json
from datetime import date, timedelta

from fastapi import HTTPException
from calendar_view import apply_calendar_status, resolve_cell
from contract_periods import day, matching_contracts, outside_contract
from planning import WEEKDAYS

from .attendance_sheet import new_sheet

PAID = 'Congé payé'
COLLECTIONS = ('vacances', 'absences', 'arret')


def packed(value):
    return json.dumps(value, sort_keys=True, separators=(',', ':'), ensure_ascii=False, default=str)


def leave_cells(row):
    current, end = day(row['start_date']), day(row['end_date'])
    while current <= end:
        for period, slot in (('matin', 'morning'), ('aprem', 'afternoon')):
            if current == day(row['start_date']) and row['start_slot'] == 'afternoon' and period == 'matin':
                continue
            if current == end and row['end_slot'] == 'morning' and period == 'aprem':
                continue
            if current == end == day(row['start_date']) and row['start_slot'] not in ('full', slot):
                continue
            yield current, period
        current += timedelta(days=1)


def cell_state(sheet, current, period):
    key = current.isoformat()
    return {'exceptions': {name: [{k: v for k, v in entry.items() if k not in ('matin', 'aprem')}
             for entry in sheet.get(name, []) if str(entry.get('date'))[:10] == key and entry.get(period)]
             for name in COLLECTIONS},
            'override': sheet.get('calendar_overrides', {}).get(key, {}).get(period),
            'schedule': copy.deepcopy(sheet.get('planning_detail', {}).get(WEEKDAYS[current.weekday()], {})),
            'resolved': list(resolve_cell(copy.deepcopy(sheet), current, period)[:2])}


def restore_cell(sheet, current, period, state):
    apply_calendar_status(sheet, current, period, 'Réinitialiser')
    key = current.isoformat()
    if state['override'] is not None:
        sheet.setdefault('calendar_overrides', {}).setdefault(key, {})[period] = copy.deepcopy(state['override'])
    for name, entries in state['exceptions'].items():
        for original in entries:
            target = next((entry for entry in sheet[name] if str(entry.get('date'))[:10] == key and
                           all(entry.get(k) == v for k, v in original.items() if k != 'date')), None)
            if target is None:
                target = {**original, 'date': key, 'matin': False, 'aprem': False}
                sheet[name].append(target)
            target[period] = True


def sheet_matches(sheet, username, current):
    return (sheet.get('account_username') == username and
            (not sheet.get('contract_id') or
             (sheet.get('sheet_month') == current.month and sheet.get('sheet_year') == current.year))
            and not outside_contract(sheet, current))


def signed_months(store, username):
    store.execute('''SELECT fiche_year,fiche_month FROM signed_fiche_documents
        WHERE employee_username=%s AND mode='production'
        AND status NOT IN ('cancelled','declined','expired','failed')''', (username,))
    return {(r['fiche_year'], r['fiche_month']) for r in store.many()}

class LeaveAttendance:
    def lock(self, store):
        # Same order as the regular attendance writer. This also serializes two
        # approvals/cancellations across different requests/workspace copies.
        if not store.sqlite:
            store.execute('SELECT username FROM users ORDER BY username FOR UPDATE')
            store.many()

    def _configs(self, store):
        suffix = '' if store.sqlite else ' FOR UPDATE'
        store.execute('SELECT user_id,form_content FROM Presence ORDER BY user_id' + suffix)
        configs = {}
        for row in store.many():
            try:
                value = row['form_content']
                config = copy.deepcopy(value) if isinstance(value, dict) else json.loads(value)
                if not isinstance(config, dict) or not isinstance(config.get('employes_data', []), list):
                    raise ValueError()
                configs[row['user_id']] = config
            except (ValueError, TypeError):
                # Never partially approve while an unreadable workspace may contain this employee.
                raise HTTPException(409, 'Une configuration de présence est illisible. Contactez un Admin.') from None
        return configs


    def prepare(self, store, row, actor, operation):
        configs = self._configs(store)
        before = copy.deepcopy(configs)
        signed = signed_months(store, row['employee_username'])
        issues, links, transfers = [], [], []
        created, skipped = 0, 0
        updated, linked = set(), set()
        cells = list(leave_cells(row))

        def issue(code, current, period='', status='', blocking=True):
            issues.append({'code': code, 'date': current.isoformat(), 'period': period,
                           'status': status, 'blocking': blocking})

        if operation == 'approve':
            account = store.user(row['employee_username'])
            if not account:
                raise HTTPException(409, 'Le compte du salarié est introuvable.')
            manager = store.user(row.get('current_manager')) if row.get('current_manager') else None
            store.execute('SELECT * FROM personnel_contracts WHERE username=%s ORDER BY start_date,id',
                          (row['employee_username'],))
            contracts = store.many()
            # Approval creates missing month-bound records in the approver's workspace.
            own = configs.setdefault(actor['username'], {'mois': cells[0][0].month,
                'annee': cells[0][0].year, 'employes_data': []})
            for year, month in sorted({(current.year, current.month) for current, _ in cells}):
                month_cells = [(d, p) for d, p in cells if d.year == year and d.month == month]
                matching = matching_contracts(contracts, month, year)
                for contract in matching:
                    if not any(day(contract['start_date']) <= d and
                               (not contract.get('end_date') or d <= day(contract['end_date'])) for d, _ in month_cells):
                        continue
                    if any(sheet_matches(sheet, row['employee_username'], current) and
                           sheet.get('contract_id') == contract['id'] for sheet in own['employes_data']
                           for current, _ in month_cells):
                        continue
                    # A signed month is never re-created as a fresh, editable copy.
                    if (year, month) in signed:
                        continue
                    record = new_sheet(account, manager, contract, month, year, own['employes_data'])
                    record['id'] = hashlib.sha256(f"leave:{row['id']}:{actor['username']}:{contract['id']}:{year}:{month}".encode()).hexdigest()[:32]
                    own['employes_data'].append(record)
                    created += 1
                for current, period in month_cells:
                    targets = [(owner, sheet) for owner, config in configs.items()
                               for sheet in config.get('employes_data', [])
                               if sheet_matches(sheet, row['employee_username'], current)]
                    if not targets:
                        issue('missing_contract', current, period)
                    for owner, sheet in targets:
                        state = cell_state(sheet, current, period)
                        status = state['resolved'][0]
                        protected = sheet.get('is_archived') or (year, month) in signed
                        if status == PAID:
                            linked.add((current.isoformat(), period))
                            changed = False
                        elif status in ('Repos', 'Férié'):
                            skipped += 1
                            continue
                        elif protected:
                            issue('protected', current, period, status)
                            continue
                        else:
                            if status != 'Travail':
                                issue('conflict', current, period, status, blocking=False)
                            apply_calendar_status(sheet, current, period, PAID)
                            updated.add((current.isoformat(), period))
                            changed = True
                        links.append({'owner': owner, 'sheet_id': str(sheet['id']), 'date': current.isoformat(),
                                      'period': period, 'before': state, 'after': cell_state(sheet, current, period),
                                      'changed': changed})
        else:
            store.execute('SELECT * FROM leave_attendance_cells WHERE request_id=%s ORDER BY id', (row['id'],))
            for record in store.many():
                if not record['changed']:
                    continue
                current, period, owner = day(record['cell_date']), record['cell_period'], record['workspace_owner']
                sheet = next((s for s in configs.get(owner, {}).get('employes_data', [])
                              if str(s.get('id')) == record['sheet_id']), None)
                if sheet is None or not sheet_matches(sheet, row['employee_username'], current):
                    issue('missing_sheet', current, period)
                    continue
                if packed(cell_state(sheet, current, period)) != record['after_state']:
                    issue('modified', current, period)
                    continue
                store.execute('''SELECT c.id FROM leave_attendance_cells c JOIN leave_requests r ON r.id=c.request_id
                    WHERE c.request_id<>%s AND c.workspace_owner=%s AND c.sheet_id=%s AND c.cell_date=%s
                    AND c.cell_period=%s AND r.status='approved' ORDER BY c.id''',
                    (row['id'], owner, record['sheet_id'], current.isoformat(), period))
                overlapping = store.one()
                if overlapping:
                    # Transfer the baseline so cancelling overlapping A never removes B's leave.
                    transfers.append((record['before_state'], record['after_state'], overlapping['id']))
                    continue
                if sheet.get('is_archived') or (current.year, current.month) in signed:
                    issue('protected', current, period, PAID)
                    continue
                restore_cell(sheet, current, period, json.loads(record['before_state']))
                updated.add((current.isoformat(), period))
        touched = {owner: config for owner, config in configs.items() if config != before.get(owner)}
        summary = {'operation': operation, 'created_sheets': created, 'updated_cells': len(updated),
                   'linked_cells': len(linked), 'skipped_cells': skipped, 'issues': issues,
                   'can_apply': not any(i['blocking'] for i in issues),
                   'needs_confirmation': any(not i['blocking'] for i in issues)}
        fingerprint = hashlib.sha256(packed({'request': row, 'before': before,
            'signed': sorted(signed), 'summary': summary, 'transfers': transfers}).encode()).hexdigest()
        return {'public': {**summary, 'fingerprint': fingerprint}, 'configs': touched,
                'links': links, 'transfers': transfers}

    def apply(self, store, row, actor, operation, confirmation=None):
        plan = self.prepare(store, row, actor, operation)
        public = plan['public']
        if not public['can_apply']:
            codes = {i['code'] for i in public['issues'] if i['blocking']}
            message = ('Complétez le contrat couvrant les dates demandées.' if 'missing_contract' in codes else
                       'La fiche est signée, en cours de signature, archivée ou modifiée. Une correction manuelle est nécessaire.')
            raise HTTPException(409, message)
        if public['needs_confirmation'] and confirmation != public['fingerprint']:
            raise HTTPException(409, 'Vérifiez les conflits de présence et confirmez la version actuelle avant d’approuver.')
        if confirmation is not None and confirmation != public['fingerprint']:
            raise HTTPException(409, 'La demande ou les présences ont changé. Vérifiez à nouveau la synchronisation.')
        for owner, config in plan['configs'].items():
            store.execute('SELECT user_id FROM Presence WHERE user_id=%s', (owner,))
            if store.one():
                store.execute('UPDATE Presence SET form_content=%s WHERE user_id=%s', (packed(config), owner))
            else:
                store.execute('INSERT INTO Presence(user_id,form_content) VALUES(%s,%s)', (owner, packed(config)))
        for record in plan['links']:
            store.execute('''INSERT INTO leave_attendance_cells
                (request_id,workspace_owner,sheet_id,cell_date,cell_period,before_state,after_state,changed)
                VALUES(%s,%s,%s,%s,%s,%s,%s,%s)''', (row['id'], record['owner'], record['sheet_id'],
                record['date'], record['period'], packed(record['before']), packed(record['after']), int(record['changed'])))
        for before_state, after_state, identifier in plan['transfers']:
            store.execute('UPDATE leave_attendance_cells SET before_state=%s,after_state=%s,changed=1 WHERE id=%s',
                          (before_state, after_state, identifier))
        store.event(row['id'], actor, 'attendance_synced', row['status'], row['status'],
                    f"{public['updated_cells']} demi-journée(s) mise(s) à jour ; {public['created_sheets']} fiche(s) créée(s).")
        return public

    def guard_signed(self, store, previous, incoming):
        old = {str(s.get('id')): s for s in previous.get('employes_data', [])}
        protected = {}
        for sheet in incoming.get('employes_data', []):
            original = old.get(str(sheet.get('id')))
            content = {k: v for k, v in sheet.items() if k != 'is_archived'}
            prior = {k: v for k, v in (original or {}).items() if k != 'is_archived'}
            if packed(content) == packed(prior) or not sheet.get('account_username'):
                continue
            username = (original or sheet).get('account_username')
            if username not in protected:
                protected[username] = signed_months(store, username)
            period = ((original or sheet).get('sheet_year'), (original or sheet).get('sheet_month'))
            # Legacy sheets have no month binding: protect them conservatively if
            # that person has any tracked signature instead of guessing a month.
            if period in protected[username] or (not (original or sheet).get('contract_id') and protected[username]):
                raise HTTPException(409, 'Cette fiche est signée ou en cours de signature. Son contenu ne peut plus être modifié.')

    def sync_new_sheets(self, store, previous, incoming, actor):
        old = {str(s.get('id')): s for s in previous.get('employes_data', [])}
        for sheet in incoming.get('employes_data', []):
            original = old.get(str(sheet.get('id')))
            if not sheet.get('account_username') or not sheet.get('contract_id') or sheet.get('is_archived'):
                continue
            if original and all(original.get(key) == sheet.get(key) for key in ('sheet_period_start', 'sheet_period_end')):
                continue
            store.execute("SELECT * FROM leave_requests WHERE employee_username=%s AND status='approved' ORDER BY id",
                          (sheet['account_username'],))
            for row in store.many():
                for current, period in leave_cells(row):
                    if not sheet_matches(sheet, row['employee_username'], current):
                        continue
                    if original and sheet_matches(original, row['employee_username'], current):
                        continue
                    store.execute('''SELECT id FROM leave_attendance_cells WHERE request_id=%s
                        AND workspace_owner=%s AND sheet_id=%s AND cell_date=%s AND cell_period=%s''',
                        (row['id'], actor['username'], str(sheet['id']), current.isoformat(), period))
                    recorded = store.one()
                    before = cell_state(sheet, current, period)
                    status = before['resolved'][0]
                    if status in ('Repos', 'Férié'):
                        continue
                    if status not in ('Travail', PAID):
                        raise HTTPException(409, 'Un congé approuvé entre en conflit avec le planning de cette nouvelle fiche.')
                    if status == 'Travail':
                        apply_calendar_status(sheet, current, period, PAID)
                    if recorded:
                        if status == 'Travail':
                            store.execute('''UPDATE leave_attendance_cells SET before_state=%s,after_state=%s,changed=1
                                WHERE id=%s''', (packed(before), packed(cell_state(sheet, current, period)), recorded['id']))
                        continue
                    store.execute('''INSERT INTO leave_attendance_cells
                        (request_id,workspace_owner,sheet_id,cell_date,cell_period,before_state,after_state,changed)
                        VALUES(%s,%s,%s,%s,%s,%s,%s,%s)''', (row['id'], actor['username'], str(sheet['id']),
                        current.isoformat(), period, packed(before), packed(cell_state(sheet, current, period)), int(status == 'Travail')))


def ensure_leave_attendance_schema(store):
    identity = 'INTEGER PRIMARY KEY AUTOINCREMENT' if store.sqlite else 'BIGINT AUTO_INCREMENT PRIMARY KEY'
    engine = '' if store.sqlite else ' ENGINE=InnoDB DEFAULT CHARSET=utf8mb4'
    store.execute(f'''CREATE TABLE IF NOT EXISTS leave_attendance_cells (
        id {identity}, request_id BIGINT NOT NULL, workspace_owner VARCHAR(255) NOT NULL,
        sheet_id VARCHAR(64) NOT NULL, cell_date DATE NOT NULL, cell_period VARCHAR(8) NOT NULL,
        before_state TEXT NOT NULL, after_state TEXT NOT NULL, changed INTEGER NOT NULL,
        UNIQUE(request_id,workspace_owner,sheet_id,cell_date,cell_period)){engine}''')
