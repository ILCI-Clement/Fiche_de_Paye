"""Limit managers to attendance changes within their team."""
import copy
from fastapi import HTTPException
from datetime import date
from calendar import monthrange
from .role_policy import is_admin, is_super_admin
from .destructive_audit import append_deletion_event, append_archive_event

ATTENDANCE_FIELDS = {'vacances', 'absences', 'arret', 'calendar_overrides', 'planning_detail', 'nb_jours', 'nb_heures_jour'}


def attendance_scope(cursor, actor, config, load_user, can_manage):
    """Project editable sheets without returning inaccessible historical content."""
    if is_admin(actor):
        return copy.deepcopy(config), []
    allowed, hidden, access = [], [], {}
    for sheet in config.get('employes_data', []):
        username = str(sheet.get('account_username') or '') if isinstance(sheet, dict) else ''
        if username not in access:
            target = load_user(cursor, username) if username else None
            access[username] = bool(target and can_manage(cursor, actor, target))
        (allowed if access[username] else hidden).append(copy.deepcopy(sheet))
    return {**copy.deepcopy(config), 'employes_data': allowed}, hidden


def preserve_hidden_sheets(previous, incoming, hidden):
    """Merge a validated visible submission into the locked complete snapshot."""
    protected = {str(sheet['id']) for sheet in hidden if isinstance(sheet, dict) and sheet.get('id') is not None}
    rows = incoming.get('employes_data', [])
    if any(str(sheet.get('id')) in protected for sheet in rows):
        raise HTTPException(403, 'Les fiches hors de votre périmètre ne peuvent pas être modifiées.')
    replacements = {str(sheet['id']): sheet for sheet in rows}
    merged = []
    # Hidden values are compared by content: scope returns independent copies.
    # Ordered traversal keeps the stored history in place and appends only new IDs.
    for sheet in previous.get('employes_data', []):
        identity = str(sheet.get('id')) if isinstance(sheet, dict) else None
        if sheet in hidden:
            merged.append(copy.deepcopy(sheet))
        elif identity in replacements:
            merged.append(replacements.pop(identity))
    merged.extend(sheet for sheet in rows if str(sheet['id']) in replacements)
    return {**copy.deepcopy(previous), **incoming, 'employes_data': merged}


def validate_manager_save(cursor, actor, previous, incoming, load_user, can_manage, *, workspace_owner=None, confirmed_deletions=None):
    admin = is_admin(actor)
    if not admin and 'Responsable' not in actor['role_tags']:
        raise HTTPException(403, 'Seul un administrateur ou responsable peut remplir les présences.')
    old = {str(e.get('id')): e for e in previous.get('employes_data', [])}
    new = incoming.get('employes_data', [])
    ids = [str(e.get('id')) for e in new]
    if len(ids) != len(set(ids)) or any(e.get('id') is None for e in new):
        raise HTTPException(400, 'Identifiants de fiches invalides.')
    removed = set(old) - set(ids)
    if removed:
        if not is_super_admin(actor):
            raise HTTPException(403, 'La suppression définitive est réservée au Super Admin. Archivez les fiches pour conserver leur historique.')
        if any(not old[key].get('is_archived') for key in removed):
            raise HTTPException(409, 'Archivez les fiches avant leur suppression définitive.')
        if removed != set(confirmed_deletions or []):
            raise HTTPException(400, 'Confirmez explicitement les numéros des fiches à supprimer définitivement.')
        for key in removed:
            append_deletion_event(cursor, actor, 'attendance_sheet',
                                  f"{workspace_owner or actor['username']}:{key}",
                                  {'deleted_records': 1, 'previous_is_archived': True})
    for employee in new:
        if type(employee.get('is_archived', False)) is not bool:
            raise HTTPException(400, 'État d’archivage invalide.')
        original = old.get(str(employee['id']))
        if admin:
            if original is not None and bool(original.get('is_archived')) != employee.get('is_archived', False):
                append_archive_event(cursor, actor, 'attendance_sheet',
                                     f"{workspace_owner or actor['username']}:{employee['id']}",
                                     employee.get('is_archived', False))
            # A new contract-bound sheet must still refer to an active contract.
            # Existing archived contracts keep their historical attendance snapshots.
            if original is None and employee.get('contract_id'):
                cursor.execute('SELECT * FROM personnel_contracts WHERE id=%s AND username=%s',
                               (employee['contract_id'], str(employee.get('account_username') or '')))
                contract = cursor.fetchone()
                if not contract or contract.get('is_archived'):
                    raise HTTPException(409, 'Le contrat est absent ou archivé. Actualisez les fiches.')
            continue
        username = str(employee.get('account_username') or '')
        target = load_user(cursor, username) if username else None
        if not target or not can_manage(cursor, actor, target):
            raise HTTPException(403, 'Cette personne ne fait pas partie de votre équipe.')
        if original:
            for field in set(original) | set(employee):
                if field not in ATTENDANCE_FIELDS and (original.get(field) or None) != (employee.get(field) or None):
                    raise HTTPException(403, 'Un responsable peut uniquement modifier les présences et le planning.')
        else:
            cursor.execute('SELECT * FROM personnel_contracts WHERE id=%s AND username=%s', (employee.get('contract_id'), username))
            contract = cursor.fetchone()
            if not contract or contract.get('is_archived'):
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
