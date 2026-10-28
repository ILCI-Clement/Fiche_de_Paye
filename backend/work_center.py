"""Read-only operational readiness and actionable work, with server-side scopes."""
from collections import Counter, defaultdict
from datetime import datetime
from zoneinfo import ZoneInfo

from fastapi import Depends, HTTPException, Query, Response
from pydantic import EmailStr, TypeAdapter, ValidationError
from contract_periods import day, matching_contracts

if __package__:
    from .role_policy import checked_resource_actor, is_admin
    from .person_names import person_display_name
    from .personnel_rules import registration_missing
    from .signed_fiche_store import signed_fiche_scope
    from .leave_rules import can_review
else:
    from role_policy import checked_resource_actor, is_admin
    from person_names import person_display_name
    from personnel_rules import registration_missing
    from signed_fiche_store import signed_fiche_scope
    from leave_rules import can_review


_EMAIL = TypeAdapter(EmailStr)
SIGNATURE_LABELS = {
    'sent': 'En attente de signature', 'partially_signed': 'Une signature reçue',
    'pending': 'Envoi à vérifier', 'uncertain': 'Envoi à vérifier',
    'declined': 'Signature refusée', 'expired': 'Demande expirée',
}


def mailbox(value):
    try:
        return str(_EMAIL.validate_python(value)).casefold()
    except (ValidationError, TypeError):
        return ''


def person_readiness(user, manager, contracts, *, actor, month, year, emails, registry=None):
    """A register warning never becomes an attendance prerequisite."""
    issues = []
    username = user['username']
    admin = is_admin(actor)
    can_setup = admin or ('Responsable' in actor.get('role_tags', [])
                          and user.get('manager_username') == actor['username'])
    def issue(code, label, domain, destination=None):
        issues.append({'code': code, 'label': label, 'domain': domain, 'destination': destination})
    contract_ready = bool(matching_contracts(contracts, month, year))
    if not contract_ready:
        issue('contract', 'Aucun contrat actif pour ce mois', 'attendance',
              {'page': 'attendance', 'username': username, 'month': month, 'year': year} if can_setup
              else {'page': 'admin', 'username': username} if admin else None)
    if not user.get('full_name') and not (user.get('first_name') and user.get('last_name')):
        issue('name', 'Nom et prénoms à compléter', 'contacts', {'page': 'admin', 'username': username} if admin else None)
    email = mailbox(user.get('email'))
    if not email or emails[email] != 1:
        issue('email', 'E-mail absent, invalide ou partagé', 'contacts', {'page': 'admin', 'username': username} if admin else None)
    manager_ok = bool(manager and manager.get('is_active', True) and not manager.get('login_deleted_at')
                      and ({'Admin', 'SuperAdmin', 'Responsable'} & set(manager.get('role_tags', [])))
                      and manager['username'] != username)
    if not manager_ok:
        issue('manager', 'Responsable direct à définir ou réactiver', 'contacts', {'page': 'admin', 'username': username} if admin else None)
    elif not mailbox(manager.get('email')) or emails[mailbox(manager.get('email'))] != 1 or not manager.get('full_name'):
        issue('manager_contact', 'Nom ou e-mail du responsable à compléter', 'contacts', {'page': 'admin', 'username': manager['username']} if admin else None)
    if admin:
        if not registry:
            issue('register_link', 'Fiche du registre à créer ou à associer', 'register', {'page': 'admin', 'username': username})
        elif registry.get('is_archived'):
            issue('register_archived', 'Fiche du registre archivée', 'register', {'page': 'personnel', 'employeeId': registry['employee_id']})
        elif registration_missing(registry, {'entry_date': registry.get('entry_date')}):
            issue('register_fields', 'Fiche du registre à compléter', 'register', {'page': 'personnel', 'employeeId': registry['employee_id']})
    return {'username': username, 'display_name': person_display_name(user),
            'attendance_ready': contract_ready,
            'contacts_ready': not any(item['domain'] == 'contacts' for item in issues),
            'issues': issues}


def register_work_center(app, *, get_db_connection, get_current_user, load_user, all_users, can_manage):
    @app.get('/dashboard/work-center')
    def work_center(response: Response, month: int | None = Query(default=None, ge=1, le=12),
                    year: int | None = Query(default=None, ge=2000, le=2100), actor=Depends(get_current_user)):
        response.headers['Cache-Control'] = 'no-store'
        today = datetime.now(ZoneInfo('Europe/Paris')).date()
        month, year = month or today.month, year or today.year
        db = get_db_connection()
        try:
            with db.cursor() as cursor:
                actor = checked_resource_actor(cursor, actor, load_user)
                if not is_admin(actor) and 'Responsable' not in actor.get('role_tags', []):
                    raise HTTPException(403, 'Accès réservé aux administrateurs et responsables.')
                accounts = all_users(cursor)
                by_name = {user['username']: user for user in accounts}
                # Same catalog as attendance, excluding inactive/deleted accounts.
                visible = [u for u in accounts if u.get('is_active', True) and not u.get('login_deleted_at')
                           and 'Employe' in u.get('role_tags', []) and (is_admin(actor) or can_manage(cursor, actor, u))]
                names = sorted(user['username'] for user in visible)
                contracts = defaultdict(list)
                if names:
                    marks = ','.join(['%s'] * len(names))
                    cursor.execute(f'''SELECT username,start_date,end_date,is_archived FROM personnel_contracts
                        WHERE username IN ({marks}) ORDER BY start_date,id''', tuple(names))
                    for row in cursor.fetchall():
                        contracts[row['username']].append(row)
                registry = {}
                unlinked = []
                if is_admin(actor):
                    # Completeness booleans only: no address, birth date, permit or NIR payload.
                    cursor.execute('''SELECT r.employee_id,r.linked_username,r.full_name,r.first_name,r.last_name,
                        r.is_archived,(r.nationality<>'') AS nationality,(r.sex<>'') AS sex,
                        (r.address<>'') AS address,(r.birth_date IS NOT NULL) AS birth_date,
                        (r.job_title<>'') AS job_title,(r.qualification<>'') AS qualification,
                        (r.personnel_type<>'') AS personnel_type,
                        (SELECT p.entry_date FROM personnel_register_periods p WHERE p.employee_id=r.employee_id
                         ORDER BY p.entry_date DESC,p.period_id DESC LIMIT 1) AS entry_date
                        FROM personnel_register r''')
                    for row in cursor.fetchall():
                        if row['linked_username']:
                            registry[row['linked_username']] = row
                        elif not row['is_archived']:
                            unlinked.append(row)
                emails = Counter(mailbox(user.get('email')) for user in accounts)
                people = [person_readiness(user, by_name.get(user.get('manager_username')), contracts[user['username']],
                          actor=actor, month=month, year=year, emails=emails, registry=registry.get(user['username']))
                          for user in visible]
                people.sort(key=lambda p: (not bool(p['issues']), p['display_name'].casefold(), p['username']))
                tasks = []
                for person in people:
                    if person['issues']:
                        tasks.append({'id': 'person:' + person['username'], 'kind': 'readiness',
                            'name': person['display_name'], 'label': ' · '.join(i['label'] for i in person['issues']),
                            'destination': next((i['destination'] for i in person['issues'] if i['destination']), None)})
                for row in unlinked:
                    tasks.append({'id': f"register:{row['employee_id']}", 'kind': 'readiness',
                        'name': person_display_name(row), 'label': 'Registre sans compte associé',
                        'destination': {'page': 'personnel', 'employeeId': row['employee_id']}})
                # Review scope is narrower than managed departments: direct manager only.
                leave_scope = '1=1' if is_admin(actor) else 'u.manager_username=%s AND r.employee_username<>%s'
                params = () if is_admin(actor) else (actor['username'], actor['username'])
                cursor.execute(f'''SELECT r.id,r.employee_username,r.start_date,r.end_date,u.manager_username
                    FROM leave_requests r LEFT JOIN users u ON u.username=r.employee_username
                    WHERE r.status='pending' AND r.is_archived=0 AND {leave_scope}
                    ORDER BY r.created_at,r.id''', params)
                for row in cursor.fetchall():
                    if can_review(actor, row['employee_username'], row.get('manager_username')):
                        tasks.append({'id': f"leave:{row['id']}", 'kind': 'leave',
                            'name': person_display_name(by_name.get(row['employee_username'])),
                            'label': f"{day(row['start_date']):%d/%m/%Y} → {day(row['end_date']):%d/%m/%Y}",
                            'destination': {'page': 'leave', 'leaveId': row['id']}})
                scope, params = signed_fiche_scope(actor)
                cursor.execute(f'''SELECT id,employee_username,status,fiche_month,fiche_year,last_error
                    FROM signed_fiche_documents WHERE mode='production' AND status<>'ready' AND status<>'cancelled'
                    AND {scope} ORDER BY created_at,id''', params)
                for row in cursor.fetchall():
                    status = row['status']
                    kind = 'archive' if status == 'completed' else 'signature'
                    label = ('Récupération du PDF à relancer' if row['last_error'] else 'PDF signé en cours de récupération') if kind == 'archive' else SIGNATURE_LABELS.get(status, 'État de signature à vérifier')
                    tasks.append({'id': 'signature:' + row['id'], 'kind': kind,
                        'name': person_display_name(by_name.get(row['employee_username'])),
                        'label': f"{row['fiche_month']:02d}/{row['fiche_year']} · {label}",
                        'destination': {'page': 'attendance', 'view': 'signed', 'month': row['fiche_month'], 'year': row['fiche_year']}})
                counts = {kind: sum(task['kind'] == kind for task in tasks) for kind in ('readiness', 'leave', 'signature', 'archive')}
                return {'month': month, 'year': year, 'people': people, 'tasks': tasks, 'counts': counts}
        finally:
            db.close()
