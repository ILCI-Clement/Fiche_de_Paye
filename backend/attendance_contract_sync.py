"""Refresh only editable month-bound contract metadata, never attendance or PDFs."""
import copy
from datetime import datetime, timezone

from contract_periods import contract_values, matching_contracts
from fiche_serialization import serialize_dates
from .leave_attendance import signed_months


def refresh_contract_snapshots(store, config, *, actor=None):
    """Without actor this is a read-only stale check; writes share the caller's lock/transaction."""
    contracts, protected, updated = {}, {}, []
    for sheet in config.get('employes_data', []):
        username, contract_id = sheet.get('account_username'), sheet.get('contract_id')
        month, year = sheet.get('sheet_month'), sheet.get('sheet_year')
        if (not username or not contract_id or sheet.get('is_archived')
                or type(month) is not int or type(year) is not int or not 1 <= month <= 12 or not 2000 <= year <= 2100):
            continue
        key = (username, contract_id)
        if key not in contracts:
            store.execute('SELECT * FROM personnel_contracts WHERE id=%s AND username=%s', (contract_id, username))
            contracts[key] = store.one()
        contract = contracts[key]
        if not contract or not matching_contracts([contract], month, year):
            # Missing/archived/nonintersecting contracts remain historical snapshots.
            continue
        values = serialize_dates(contract_values(contract, month, year))
        changed = any(serialize_dates(sheet.get(k)) != value for k, value in values.items() if k != 'contract_snapshot')
        before = sheet.get('contract_snapshot') or {}
        changed = changed or any(str(before.get(k)) != str(values['contract_snapshot'].get(k))
                                for k in ('id', 'username', 'employee_type', 'start_date', 'end_date'))
        changed = changed or bool(before.get('is_cdi')) != bool(contract['is_cdi'])
        if not changed:
            continue
        if username not in protected:
            protected[username] = signed_months(store, username)
        if (year, month) in protected[username]:
            continue
        updated.append(str(sheet['id']))
        if actor is not None:
            sheet.setdefault('contract_sync_history', []).append({
                'at': datetime.now(timezone.utc).isoformat(), 'actor': actor['username'],
                'before': {k: copy.deepcopy(sheet.get(k)) for k in values},
            })
            sheet.update(values)
    return updated
