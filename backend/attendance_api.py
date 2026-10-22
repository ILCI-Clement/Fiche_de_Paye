"""React attendance adapters around the existing storage, calendar and generators."""
import copy
import io
import json
import zipfile
from calendar import monthrange
from datetime import date
from typing import Literal
from urllib.parse import quote

from fastapi import Depends, HTTPException, Query
from fastapi.responses import Response
from pydantic import BaseModel, Field

from contract_periods import bind_contract, matching_contracts, outside_contract, period_sheets
from calendar_view import apply_calendar_status, resolve_cell
from fiche_exports import generate_fiche_file, validate_fiche
from fiche_serialization import deserialize_dates, serialize_dates
from planning import default_planning
from .attendance_revision import config_revision
from .role_policy import is_admin, is_super_admin
from .person_names import load_person_names, person_display_name


class OpenSheet(BaseModel):
    month: int = Field(ge=1, le=12)
    year: int = Field(ge=2000, le=2100)
    revision: str | None = Field(default=None, min_length=64, max_length=64)


class SaveWorkspace(BaseModel):
    config: dict
    revision: str = Field(min_length=64, max_length=64)
    deleted_ids: list[str] = Field(default_factory=list, max_length=100)


class ExportSheets(OpenSheet):
    ids: list[str] = Field(min_length=1, max_length=100)
    format: Literal['original', 'pdf'] = 'original'


class CalendarCell(BaseModel):
    date: date
    period: Literal['matin', 'aprem']


class CalendarEdit(OpenSheet):
    cells: list[CalendarCell] = Field(min_length=1, max_length=62)
    status: Literal['Travail', 'Congé payé', 'Absence', 'Arrêt maladie', 'Férié', 'Autre', 'Repos', 'Réinitialiser']
    reason: str = Field(default='', max_length=2000)


def register_attendance_routes(app, *, get_db_connection, get_current_user, load_user, can_manage, save_config, save_checked=None):
    def require_role(actor):
        if not is_admin(actor) and 'Responsable' not in actor['role_tags']:
            raise HTTPException(403, 'La saisie des présences est réservée aux administrateurs et responsables.')

    def read_snapshot(cursor, actor):
        cursor.execute('SELECT form_content FROM Presence WHERE user_id=%s', (actor['username'],))
        row = cursor.fetchone()
        if not row:
            return {'mois': date.today().month, 'annee': date.today().year, 'employes_data': []}, config_revision({})
        try:
            value = row['form_content']
            config = value if isinstance(value, dict) else json.loads(value)
            if not isinstance(config, dict) or not isinstance(config.get('employes_data', []), list):
                raise ValueError()
            return config, config_revision(config)
        except (TypeError, ValueError):
            raise HTTPException(422, 'Configuration enregistrée invalide. Contactez un administrateur.') from None

    def read_config(cursor, actor):
        return read_snapshot(cursor, actor)[0]

    def commit(actor, config, revision, deleted_ids=None):
        if save_checked:
            if deleted_ids:
                return save_checked(actor['username'], serialize_dates(config), actor, revision,
                                    confirmed_deletions=deleted_ids)
            return save_checked(actor['username'], serialize_dates(config), actor, revision)
        save_config(actor['username'], serialize_dates(config), actor)
        return {'config': serialize_dates(config), 'revision': config_revision(config)}

    def check_client_revision(stored_revision, supplied):
        if save_checked and supplied is None:
            raise HTTPException(428, 'Rechargez la page pour obtenir la version actuelle de la configuration.')
        if supplied is not None and stored_revision != supplied:
            raise HTTPException(409, 'La configuration a changé. Votre brouillon est conservé ; rechargez le serveur.')

    @app.post('/attendance/save')
    def save_workspace(payload: SaveWorkspace, actor=Depends(get_current_user)):
        require_role(actor)
        return commit(actor, payload.config, payload.revision, payload.deleted_ids)

    def require_employee(cursor, actor, employee):
        if is_admin(actor):
            return
        target = load_user(cursor, str(employee.get('account_username') or ''))
        if not target or not can_manage(cursor, actor, target):
            raise HTTPException(403, 'Cette fiche ne fait pas partie de votre équipe.')

    def selected(config, sheet_id):
        result = next((e for e in config.get('employes_data', []) if str(e.get('id')) == sheet_id), None)
        if result is None:
            raise HTTPException(404, 'Fiche introuvable.')
        return result

    @app.get('/attendance/workspace')
    def workspace(actor=Depends(get_current_user)):
        require_role(actor)
        db = get_db_connection()
        try:
            with db.cursor() as cursor:
                config, revision = read_snapshot(cursor, actor)
                cursor.execute('SELECT username FROM users ORDER BY username')
                users = [load_user(cursor, row['username']) for row in cursor.fetchall()]
                names = load_person_names(cursor)
                accounts = {user['username']: user for user in users if user}
                cursor.execute('SELECT id,name FROM organization_groups WHERE is_active=1 ORDER BY name')
                group_names = {row['id']: row['name'] for row in cursor.fetchall()}
                catalog = []
                for user in users:
                    if user and user.get('is_active', True) and ('Admin' in actor['role_tags'] or can_manage(cursor, actor, user)):
                        cursor.execute('SELECT * FROM personnel_contracts WHERE username=%s ORDER BY start_date,id', (user['username'],))
                        manager = next((u for u in users if u and u['username'] == user.get('manager_username')), None)
                        catalog.append({'username': user['username'], 'email': user['email'],
                            'display_name': person_display_name(user, names),
                            'group_names': [group_names[group_id] for group_id in user.get('group_ids', []) if group_id in group_names],
                            'manager_username': user.get('manager_username'), 'manager_email': manager['email'] if manager else '',
                            'manager_display_name': person_display_name(manager, names) if manager else 'Aucun responsable',
                            'contracts': cursor.fetchall()})
                # Reject unauthorized legacy records without exposing them in the new UI.
                visible = []
                for employee in config.get('employes_data', []):
                    try:
                        require_employee(cursor, actor, employee)
                        visible.append(employee)
                    except HTTPException:
                        pass
                if 'Admin' not in actor['role_tags'] and len(visible) != len(config.get('employes_data', [])):
                    # Never save a filtered configuration, which would silently delete hidden history.
                    raise HTTPException(409, 'Cette configuration contient des fiches hors de votre équipe. Un Admin doit la vérifier avant modification.')
                return {'config': config, 'revision': revision, 'accounts': catalog,
                        'person_labels': {str(employee['id']): person_display_name(
                            accounts.get(str(employee.get('account_username') or '')) or
                            {'email': employee.get('email_employe')}, names)
                            for employee in visible if employee.get('id') is not None},
                        'can_edit_identity': is_admin(actor), 'can_archive_sheets': is_admin(actor),
                        'can_delete_sheets': is_super_admin(actor)}
        finally:
            db.close()

    @app.post('/attendance/open/{username}')
    def open_sheet(username: str, payload: OpenSheet, actor=Depends(get_current_user)):
        require_role(actor)
        db = get_db_connection()
        try:
            with db.cursor() as cursor:
                account = load_user(cursor, username)
                if not account or not can_manage(cursor, actor, account):
                    raise HTTPException(403, 'Cette personne ne vous est pas accessible.')
                config, revision = read_snapshot(cursor, actor)
                check_client_revision(revision, payload.revision)
                cursor.execute('SELECT * FROM personnel_contracts WHERE username=%s ORDER BY start_date,id', (username,))
                contracts = matching_contracts(cursor.fetchall(), payload.month, payload.year)
                records = config.setdefault('employes_data', [])
                saved_period = [e for e in records if e.get('account_username') == account['username']
                                and e.get('sheet_month') == payload.month and e.get('sheet_year') == payload.year]
                active = None
                if not contracts:
                    # A later contract edit or archive must not make an existing
                    # month-bound snapshot inaccessible. Only creation needs a
                    # currently matching contract; keep stored history intact.
                    existing = next((e for e in saved_period if not e.get('is_archived')), None)
                    if existing:
                        active = str(existing['id'])
                    elif saved_period:
                        raise HTTPException(409, 'Cette fiche est archivée. Un administrateur peut la restaurer depuis les archives.')
                    else:
                        raise HTTPException(409, 'Aucun contrat actif ne couvre ce mois. Vérifiez les dates du contrat dans Administration avant de créer une nouvelle fiche. Les fiches déjà enregistrées pour les autres mois sont conservées.')
                manager = load_user(cursor, account['manager_username']) if account.get('manager_username') else None
                for contract in contracts:
                    existing = next((e for e in records if e.get('account_username') == username and e.get('contract_id') == contract['id']
                        and e.get('sheet_month') == payload.month and e.get('sheet_year') == payload.year), None)
                    if existing:
                        if existing.get('is_archived'):
                            raise HTTPException(409, 'Cette fiche est archivée. Un administrateur peut la restaurer depuis les archives.')
                        active = str(existing['id'])
                        continue
                    name = person_display_name(account)
                    record = {'nom': name, 'nom_stagiaire': name, 'prenom_stagiaire': '',
                        'email_employe': account['email'], 'account_username': username,
                        'responsable': person_display_name(manager) if manager else '',
                        'email_responsable': (manager or {}).get('email', ''),
                        'nb_jours': 0, 'nb_heures_jour': 0.0, 'taux_horaire': 0.0,
                        'facture_mensuelle': 0.0, 'transport': '', 'taux': 0.0,
                        'vacances': [], 'absences': [], 'arret': [], 'calendar_overrides': {}, 'planning_detail': default_planning()}
                    prior = next((e for e in reversed(records) if e.get('account_username') == username), None)
                    if prior:
                        for field in ('nom', 'nom_stagiaire', 'prenom_stagiaire', 'responsable', 'email_responsable', 'planning_detail'):
                            if prior.get(field):
                                record[field] = copy.deepcopy(prior[field])
                        expected_type = 'Stagiaire' if contract['employee_type'] == 'stagiaire' else 'Salarié'
                        if prior.get('type') == expected_type:
                            for field in ('taux_horaire', 'nb_heures_jour', 'transport', 'taux'):
                                if field in prior:
                                    record[field] = copy.deepcopy(prior[field])
                    records.append(serialize_dates(bind_contract(record, contract, payload.month, payload.year)))
                    active = str(records[-1]['id'])
                config.update(mois=payload.month, annee=payload.year)
            result = commit(actor, config, revision)
            return {**result, 'active_id': active}
        finally:
            db.close()

    @app.get('/attendance/calendar/{sheet_id}')
    def calendar_data(sheet_id: str, month: int = Query(ge=1, le=12), year: int = Query(ge=2000, le=2100), actor=Depends(get_current_user)):
        require_role(actor)
        db = get_db_connection()
        try:
            with db.cursor() as cursor:
                employee = selected(read_config(cursor, actor), sheet_id)
                require_employee(cursor, actor, employee)
            days = []
            for number in range(1, monthrange(year, month)[1] + 1):
                current = date(year, month, number)
                days.append({'date': current.isoformat(), 'number': number, 'weekday': current.weekday(),
                    'disabled': outside_contract(employee, current),
                    'matin': resolve_cell(employee, current, 'matin')[:2], 'aprem': resolve_cell(employee, current, 'aprem')[:2]})
            return {'days': days}
        finally:
            db.close()

    @app.post('/attendance/calendar/{sheet_id}')
    def edit_calendar(sheet_id: str, payload: CalendarEdit, actor=Depends(get_current_user)):
        require_role(actor)
        db = get_db_connection()
        try:
            with db.cursor() as cursor:
                config, revision = read_snapshot(cursor, actor)
                check_client_revision(revision, payload.revision)
                employee = selected(config, sheet_id)
                require_employee(cursor, actor, employee)
            for cell in payload.cells:
                if cell.date.month != payload.month or cell.date.year != payload.year:
                    raise HTTPException(400, 'La sélection doit rester dans le mois affiché.')
                apply_calendar_status(employee, cell.date, cell.period, payload.status, payload.reason)
            config.update(mois=payload.month, annee=payload.year)
            return commit(actor, config, revision)
        finally:
            db.close()

    @app.post('/attendance/export')
    def export(payload: ExportSheets, actor=Depends(get_current_user)):
        require_role(actor)
        db = get_db_connection()
        try:
            with db.cursor() as cursor:
                config = read_config(cursor, actor)
                records = []
                seen = set()
                visible = period_sheets(config.get('employes_data', []), payload.month, payload.year)
                for sheet_id in payload.ids:
                    employee = selected(config, sheet_id)
                    if employee not in visible:
                        raise HTTPException(400, 'Une fiche sélectionnée appartient à un autre mois.')
                    require_employee(cursor, actor, employee)
                    if sheet_id not in seen:
                        records.append(employee); seen.add(sheet_id)
            generated = []
            from .esign import convert_fiche_to_pdf
            for index, employee in enumerate(records):
                restored = deserialize_dates(copy.deepcopy(employee))
                try:
                    validate_fiche(restored, require_first_name=restored.get('type') == 'Stagiaire', check_exceptions=True)
                    item = generate_fiche_file(restored, payload.month, payload.year, ordinal=index + 1 if len(records) > 1 else None)
                except (ValueError, TypeError, KeyError) as error:
                    raise HTTPException(422, str(error)) from None
                if payload.format == 'pdf':
                    item.update(data=convert_fiche_to_pdf(item['data'], item['filename']), mime='application/pdf')
                    item['filename'] = item['filename'].rsplit('.', 1)[0] + '.pdf'
                generated.append(item)
            if len(generated) == 1:
                content, filename, mime = generated[0]['data'], generated[0]['filename'], generated[0]['mime']
            else:
                result = io.BytesIO()
                with zipfile.ZipFile(result, 'w', zipfile.ZIP_DEFLATED) as archive:
                    for item in generated:
                        archive.writestr(item['filename'], item['data'])
                content, filename, mime = result.getvalue(), f'presence_{payload.year}_{payload.month:02d}.zip', 'application/zip'
            return Response(content, media_type=mime, headers={'Content-Disposition': "attachment; filename*=UTF-8''" + quote(filename)})
        finally:
            db.close()
