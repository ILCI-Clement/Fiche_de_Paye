"""Authenticated paid leave routes; no attendance, mail or balance side effects."""
from typing import Literal

from fastapi import Depends, HTTPException, Query
from pydantic import EmailStr, TypeAdapter, ValidationError
if __package__:
    from .person_names import person_display_name
else:
    from person_names import person_display_name

if __package__:
    from .leave_rules import LeaveInput, LeaveDecision, can_review, display_name, roles
    from .leave_store import LeaveStore, ensure_leave_schema, iso, timestamp
else:
    from leave_rules import LeaveInput, LeaveDecision, can_review, display_name, roles
    from leave_store import LeaveStore, ensure_leave_schema, iso, timestamp


_HISTORICAL_EMAIL = TypeAdapter(EmailStr)


def historical_actor_name(event, names):
    """Use a current public label or an already-recorded name, never a raw ID.

    An account rename intentionally preserves audit identity spelling. Its old
    name snapshot can still be displayed without guessing a new association.
    """
    username = event.get('actor_username')
    if username in names:
        return names[username]
    recorded = str(event.get('actor_name') or '').strip()
    if not recorded or recorded == str(username or '').strip():
        return 'Ancien compte'
    if '@' in recorded:
        try:
            email = str(_HISTORICAL_EMAIL.validate_python(recorded))
        except ValidationError:
            return 'Ancien compte'
        # EmailStr also accepts display-name wrappers; only a whole mailbox is
        # safe as this fallback, with its domain normalization permitted.
        if email.casefold() != recorded.casefold():
            return 'Ancien compte'
        return person_display_name({'email': email})
    return recorded


def projection(row, actor, names=None):
    current_manager = row.get('current_manager')
    return {
        **{key: iso(row[key]) for key in ('id', 'employee_username', 'start_date', 'end_date',
           'start_slot', 'end_slot', 'note', 'status', 'created_at', 'updated_at')},
        'employee_name': person_display_name({'username': row.get('employee_username'), 'email': row.get('current_email')}, names),
        'approver_name': person_display_name({'username': current_manager, 'email': row.get('manager_email')}, names) if current_manager else 'Administration',
        'is_archived': bool(row.get('is_archived', False)),
        'can_review': not row.get('is_archived') and row['status'] == 'pending' and can_review(actor, row['employee_username'], current_manager),
        'can_withdraw': not row.get('is_archived') and row['status'] == 'pending' and row['employee_username'] == actor['username'],
    }


def require_visible(row, actor):
    if not row:
        raise HTTPException(404, 'Demande introuvable.')
    if row['employee_username'] != actor['username'] and not can_review(actor, row['employee_username'], row.get('current_manager')):
        raise HTTPException(403, 'Cette demande ne vous est pas accessible.')


def register_leave_routes(app, *, get_db_connection, get_current_user):
    @app.on_event('startup')
    def migrate_leave():
        ensure_leave_schema(get_db_connection)

    @app.get('/leave/session')
    def session(actor=Depends(get_current_user)):
        with LeaveStore(get_db_connection) as store:
            user = store.user(actor['username'])
            if not user:
                raise HTTPException(401, 'Session invalide.')
            manager = store.user(user['manager_username']) if user.get('manager_username') else None
        return {'username': actor['username'], 'name': display_name(user),
                'role_tags': sorted(roles(actor)),
                'approver_name': display_name(manager) if manager else 'Administration',
                'can_review': bool({'Admin', 'Responsable'} & roles(actor))}

    @app.get('/leave-requests')
    def list_requests(scope: Literal['mine', 'team'] = 'mine',
                      status: Literal['all', 'pending', 'approved', 'rejected'] = 'all',
                      page: int = Query(default=1, ge=1),
                      page_size: int = Query(default=12, ge=1, le=50),
                      actor=Depends(get_current_user)):
        condition = "r.status <> 'withdrawn' AND r.is_archived=0"
        params = []
        if scope == 'mine':
            condition += ' AND r.employee_username=%s'
            params.append(actor['username'])
        elif 'Admin' not in roles(actor):
            if 'Responsable' not in roles(actor):
                raise HTTPException(403, 'Accès réservé aux responsables.')
            condition += ' AND u.manager_username=%s AND r.employee_username<>%s'
            params += [actor['username'], actor['username']]
        # Scope comes from the authenticated account and the current users relation.
        joins = '''FROM leave_requests r LEFT JOIN users u ON u.username=r.employee_username
                   LEFT JOIN users m ON m.username=u.manager_username'''
        with LeaveStore(get_db_connection) as store:
            store.execute(f'''SELECT COUNT(*) AS total,
                COALESCE(SUM(CASE WHEN r.status='pending' THEN 1 ELSE 0 END),0) AS pending,
                COALESCE(SUM(CASE WHEN r.status='approved' THEN 1 ELSE 0 END),0) AS approved,
                COALESCE(SUM(CASE WHEN r.status='rejected' THEN 1 ELSE 0 END),0) AS rejected
                {joins} WHERE {condition}''', tuple(params))
            counts = store.one()
            if status != 'all':
                condition += ' AND r.status=%s'
                params.append(status)
            store.execute(f'SELECT COUNT(*) AS total {joins} WHERE {condition}', tuple(params))
            total = store.one()['total']
            store.execute(f'''SELECT r.*,u.email AS current_email,m.email AS manager_email,
                m.username AS current_manager {joins} WHERE {condition}
                ORDER BY r.id DESC LIMIT %s OFFSET %s''', (*params, page_size, (page - 1) * page_size))
            rows = store.many()
            names = store.person_names()
        return {'items': [projection(row, actor, names) for row in rows], 'total': total,
                'page': page, 'page_size': page_size, 'counts': counts}

    @app.post('/leave-requests', status_code=201)
    def create_request(payload: LeaveInput, actor=Depends(get_current_user)):
        with LeaveStore(get_db_connection, write=True) as store:
            user = store.user(actor['username'])
            if not user:
                raise HTTPException(401, 'Session invalide.')
            now = timestamp()
            store.execute('''INSERT INTO leave_requests
                (employee_username,employee_name,start_date,end_date,start_slot,end_slot,note,status,created_at,updated_at)
                VALUES (%s,%s,%s,%s,%s,%s,%s,'pending',%s,%s)''',
                (actor['username'], display_name(user), payload.start_date.isoformat(), payload.end_date.isoformat(),
                 payload.start_slot, payload.end_slot, payload.note, now, now))
            request_id = store.cursor.lastrowid
            store.event(request_id, actor, 'submitted', None, 'pending', payload.note)
            return projection(store.request(request_id), actor, store.person_names())

    @app.get('/leave-requests/{request_id}')
    def detail(request_id: int, actor=Depends(get_current_user)):
        with LeaveStore(get_db_connection) as store:
            row = store.request(request_id)
            require_visible(row, actor)
            store.execute('''SELECT id,actor_name,actor_username,action,previous_status,new_status,comment,created_at
                FROM leave_events WHERE request_id=%s ORDER BY id''', (request_id,))
            history = store.many()
            actor_names = store.display_accounts(event['actor_username'] for event in history)
            for event in history:
                event['actor_name'] = historical_actor_name(event, actor_names)
            names = store.person_names()
        return {'request': projection(row, actor, names), 'history': history}

    def transition(request_id, actor, new_status, comment=''):
        with LeaveStore(get_db_connection, write=True) as store:
            row = store.request(request_id, lock=True)
            require_visible(row, actor)
            if row.get('is_archived'):
                raise HTTPException(409, 'Cette demande est archivée. Restaurez-la avant de la modifier.')
            if new_status == 'withdrawn':
                if row['employee_username'] != actor['username']:
                    raise HTTPException(403, 'Vous pouvez uniquement retirer vos propres demandes.')
            elif not can_review(actor, row['employee_username'], row.get('current_manager')):
                raise HTTPException(403, 'Seul le responsable actuel ou un Admin peut traiter cette demande.')
            if row['status'] != 'pending':
                raise HTTPException(409, 'Cette demande a déjà été traitée ou retirée. Actualisez la liste.')
            store.execute("UPDATE leave_requests SET status=%s,updated_at=%s WHERE id=%s AND status='pending'",
                          (new_status, timestamp(), request_id))
            if store.cursor.rowcount != 1:
                raise HTTPException(409, 'Cette demande vient de changer. Actualisez la liste.')
            action = 'withdrawn' if new_status == 'withdrawn' else 'decided'
            store.event(request_id, actor, action, 'pending', new_status, comment.strip())
            return projection(store.request(request_id), actor, store.person_names())

    @app.post('/leave-requests/{request_id}/withdraw')
    def withdraw(request_id: int, actor=Depends(get_current_user)):
        return transition(request_id, actor, 'withdrawn')

    @app.post('/leave-requests/{request_id}/decision')
    def decide(request_id: int, payload: LeaveDecision, actor=Depends(get_current_user)):
        return transition(request_id, actor, payload.decision, payload.comment)
