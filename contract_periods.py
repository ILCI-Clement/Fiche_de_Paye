"""Match saved contracts to a month without replacing historic sheet data."""
from calendar import monthrange
from datetime import date
from uuid import uuid4


def day(value):
    return value if isinstance(value, date) else date.fromisoformat(str(value))


def matching_contracts(contracts, month, year):
    first = date(year, month, 1)
    last = date(year, month, monthrange(year, month)[1])
    return [c for c in contracts if not c.get('is_archived') and day(c['start_date']) <= last
            and (not c.get('end_date') or day(c['end_date']) >= first)]


def bind_contract(fiche, contract, month, year):
    start = day(contract['start_date'])
    end = day(contract['end_date']) if contract.get('end_date') else None
    fiche.update(contract_id=contract['id'], contract_snapshot=dict(contract), sheet_month=month, sheet_year=year,
                 sheet_period_start=max(start,date(year,month,1)).isoformat(),
                 sheet_period_end=min(end or date.max,date(year,month,monthrange(year,month)[1])).isoformat())
    fiche['id'] = uuid4().hex
    intern = contract['employee_type'] == 'stagiaire'
    fiche['type'] = 'Stagiaire' if intern else 'Salarié'
    fiche['dds' if intern else 'ddc'] = start
    fiche['fds' if intern else 'fdc'] = end if end else 'Pas de fin'
    fiche['cdi'] = bool(contract['is_cdi']) if not intern else False
    return fiche


def outside_contract(employee, current):
    return bool(employee.get('contract_id') and
                (current < day(employee['sheet_period_start']) or current > day(employee['sheet_period_end'])))


def period_sheets(employees, month, year):
    return [e for e in employees if not e.get('contract_id') or
            (e.get('sheet_month') == month and e.get('sheet_year') == year)]


def contract_filename(name, employee):
    if not employee.get('contract_id'):
        return name
    return f"{name}_contrat_{employee['contract_id']}_{employee['sheet_period_start']}_{employee['sheet_period_end']}"


def monthly_exceptions(employee, month, year):
    """Filter export data without deleting records saved for other months."""
    result = dict(employee)
    first = date(year, month, 1)
    last = date(year, month, monthrange(year, month)[1])
    if employee.get('contract_id'):
        first = max(first, day(employee['sheet_period_start']))
        last = min(last, day(employee['sheet_period_end']))
    for collection in ('vacances', 'absences', 'arret'):
        filtered = []
        for entry in employee.get(collection, []):
            try:
                current = day(entry.get('date'))
            except (ValueError, TypeError):
                continue
            if first <= current <= last:
                filtered.append(entry)
        result[collection] = filtered
    return result
