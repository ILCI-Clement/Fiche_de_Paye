"""Limit managers to attendance changes within their team."""
from fastapi import HTTPException
from datetime import date
from calendar import monthrange

ATTENDANCE_FIELDS = {'vacances', 'absences', 'arret', 'calendar_overrides', 'planning_detail', 'nb_jours', 'nb_heures_jour'}


def validate_manager_save(cursor, actor, previous, incoming, load_user, can_manage):
    if 'Admin' in actor['role_tags']:
        return
    if 'Responsable' not in actor['role_tags']:
        raise HTTPException(403, 'Seul un administrateur ou responsable peut remplir les présences.')
    old = {str(e.get('id')): e for e in previous.get('employes_data', [])}
    new = incoming.get('employes_data', [])
    ids = [str(e.get('id')) for e in new]
    if len(ids) != len(set(ids)) or any(e.get('id') is None for e in new):
        raise HTTPException(400, 'Identifiants de fiches invalides.')
    if not set(old).issubset(ids):
        raise HTTPException(403, 'Un responsable ne peut pas supprimer une fiche.')
    for employee in new:
        username = str(employee.get('account_username') or '')
        target = load_user(cursor, username) if username else None
        if not target or not can_manage(cursor, actor, target):
            raise HTTPException(403, 'Cette personne ne fait pas partie de votre équipe.')
        original = old.get(str(employee['id']))
        if original:
            for field in set(original) | set(employee):
                if field not in ATTENDANCE_FIELDS and (original.get(field) or None) != (employee.get(field) or None):
                    raise HTTPException(403, 'Un responsable peut uniquement modifier les présences et le planning.')
        else:
            cursor.execute('SELECT * FROM personnel_contracts WHERE id=%s AND username=%s', (employee.get('contract_id'), username))
            contract = cursor.fetchone()
            if not contract:
                raise HTTPException(403, 'Un contrat enregistré est requis pour créer une fiche.')
            try:
                month, year = int(employee['sheet_month']), int(employee['sheet_year'])
                first, last = date(year, month, 1), date(year, month, monthrange(year, month)[1])
                start = max(first, contract['start_date'])
                end = min(last, contract['end_date'] or date.max)
            except (ValueError, KeyError, TypeError):
                raise HTTPException(400, 'Période de fiche invalide.')
            intern = contract['employee_type'] == 'stagiaire'
            expected = {'sheet_period_start': str(start), 'sheet_period_end': str(end),
                        'type': 'Stagiaire' if intern else 'Salarié',
                        'dds' if intern else 'ddc': str(contract['start_date']),
                        'fds' if intern else 'fdc': str(contract['end_date']) if contract['end_date'] else 'Pas de fin'}
            if start > end or any(str(employee.get(k)) != v for k,v in expected.items()):
                raise HTTPException(403, 'La fiche doit correspondre aux dates du contrat enregistré.')
            snapshot = employee.get('contract_snapshot', {})
            for field in ('id','employee_type','start_date','end_date','is_cdi'):
                if str(snapshot.get(field)) != str(contract.get(field)):
                    # JSON represents database boolean values as bool instead of int.
                    if field != 'is_cdi' or bool(snapshot.get(field)) != bool(contract.get(field)):
                        raise HTTPException(403, 'Les informations du contrat ne peuvent pas être modifiées.')
