"""Admin-only employee register, independent from login accounts and contract management.

Every mutation locks users before the employee and checks the actor again. Ordinary
changes retain snapshots; confirmed SuperAdmin deletion removes archived registry PII
and retains only the separate minimal permanent-deletion ledger.
"""
import json
import logging
import os

from fastapi import APIRouter, Depends, HTTPException, Path, Query
from fastapi.exceptions import RequestValidationError
from fastapi.exception_handlers import http_exception_handler, request_validation_exception_handler
from fastapi.routing import APIRoute
from fastapi.responses import JSONResponse
from pydantic import BaseModel, ConfigDict, Field
from pymysql.err import IntegrityError, OperationalError

try:
    from .personnel_rules import (
        ArchiveFilter, ArchiveMutation, EmployeeCreate, EmployeeUpdate,
        EmploymentStatus, MAX_IDENTIFIER, PeriodMutation, PersonnelType, like_search, paris_today,
    )
    from .personnel_store import (
        EMPLOYEE_COLUMNS, EMPLOYEE_FIELDS, PERIOD_COLUMNS, PERIOD_FIELDS,
        append_personnel_event, delete_personnel_records, employee_record, ensure_personnel_schema, insert_period,
        load_employee_detail, period_record, rename_personnel_account_link, selected_columns,
        sensitive_metadata, touch_employee, utc_timestamp,
    )
    from .personnel_sensitive import (
        NirProtectionError, decrypt_social_security_number, encrypt_social_security_number, nir_storage_ready,
    )
    from .destructive_audit import append_deletion_event
    from .role_policy import checked_resource_actor, is_admin, is_super_admin
    from .person_names import load_display_names, person_display_name
except ImportError:
    from personnel_rules import (
        ArchiveFilter, ArchiveMutation, EmployeeCreate, EmployeeUpdate,
        EmploymentStatus, MAX_IDENTIFIER, PeriodMutation, PersonnelType, like_search, paris_today,
    )
    from personnel_store import (
        EMPLOYEE_COLUMNS, EMPLOYEE_FIELDS, PERIOD_COLUMNS, PERIOD_FIELDS,
        append_personnel_event, delete_personnel_records, employee_record, ensure_personnel_schema, insert_period,
        load_employee_detail, period_record, rename_personnel_account_link, selected_columns,
        sensitive_metadata, touch_employee, utc_timestamp,
    )
    from personnel_sensitive import (
        NirProtectionError, decrypt_social_security_number, encrypt_social_security_number, nir_storage_ready,
    )
    from destructive_audit import append_deletion_event
    from role_policy import checked_resource_actor, is_admin, is_super_admin
    from person_names import load_display_names, person_display_name


_SAFE_ERROR_LOCATIONS = frozenset({
    'body', 'query', 'path', 'header', 'cookie', 'employee_id', 'period_id', 'initial_period',
    'expected_version', 'confirmation', 'is_archived', 'social_security_number', 'q', 'limit', 'page', 'page_size',
    'archive', 'status', *EMPLOYEE_FIELDS, *PERIOD_FIELDS,
    'username', 'expected_employee_id', 'target_version',
})
_LATEST_PERIOD_JOIN = '''LEFT JOIN personnel_register_periods p ON p.period_id=(
    SELECT p2.period_id FROM personnel_register_periods p2
    WHERE p2.employee_id=e.employee_id ORDER BY p2.entry_date DESC,p2.period_id DESC LIMIT 1)'''
_STATUS_FILTER = '''CASE WHEN p.entry_date IS NULL THEN 'unknown'
    WHEN p.entry_date>%s THEN 'upcoming'
    WHEN p.exit_date IS NOT NULL AND p.exit_date<%s THEN 'left'
    ELSE 'employed' END'''


class PersonnelDeleteInput(BaseModel):
    model_config = ConfigDict(extra='forbid', str_strip_whitespace=True)
    expected_version: int = Field(strict=True, ge=1, le=2_147_483_646)
    confirmation: str = Field(strict=True, min_length=1, max_length=200)


class AccountPersonnelLinkInput(BaseModel):
    model_config = ConfigDict(extra='forbid', hide_input_in_errors=True)
    employee_id: int | None = Field(ge=1, le=MAX_IDENTIFIER, strict=True)
    expected_employee_id: int | None = Field(ge=1, le=MAX_IDENTIFIER, strict=True)
    expected_version: int | None = Field(ge=1, le=2_147_483_646, strict=True)
    target_version: int | None = Field(ge=1, le=2_147_483_646, strict=True)


class PersonnelRoute(APIRoute):
    """No cacheable personnel response or echoed rejected personal input."""
    def get_route_handler(self):
        original = super().get_route_handler()

        async def handler(request):
            try:
                response = await original(request)
            except RequestValidationError as error:
                safe_errors = [{
                    'type': item['type'], 'msg': item['msg'],
                    'loc': tuple(part if part in _SAFE_ERROR_LOCATIONS else 'field'
                                 for part in item.get('loc', ())),
                } for item in error.errors()]
                response = await request_validation_exception_handler(request, RequestValidationError(safe_errors))
            except HTTPException as error:
                response = await http_exception_handler(request, error)
            except Exception:
                # Do not log request bodies, SQL errors or exception representations containing PII.
                logging.getLogger(__name__).error('Personnel register request failed.')
                response = JSONResponse(status_code=500, content={'detail': 'Le registre est temporairement indisponible.'})
            response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, private'
            response.headers['Pragma'] = 'no-cache'
            response.headers['Expires'] = '0'
            return response
        return handler


def register_personnel_routes(app, *, get_db_connection, get_current_user, load_user):
    router = APIRouter(route_class=PersonnelRoute)

    def nir_enabled():
        return os.getenv('PRESENCE_PERSONNEL_NIR_ENABLED') == '1' and nir_storage_ready()

    def require_nir_storage():
        if not nir_enabled():
            raise HTTPException(503, 'La protection des données sensibles est temporairement indisponible.')

    def checked_admin(cursor, actor, *, lock=False):
        current = checked_resource_actor(cursor, actor, load_user, lock=lock)
        if not is_admin(current):
            raise HTTPException(403, 'Le registre du personnel est réservé aux administrateurs actifs.')
        return current

    def detail_or_404(cursor, employee_id, *, lock=False):
        detail = load_employee_detail(cursor, employee_id, lock=lock)
        if not detail:
            raise HTTPException(404, 'Salarié introuvable.')
        return detail

    def editable_detail(cursor, employee_id, expected_version):
        detail = detail_or_404(cursor, employee_id, lock=True)
        if detail['employee']['version'] != expected_version:
            raise HTTPException(409, 'Cette fiche a changé. Actualisez-la avant de réessayer.')
        return detail

    def canonical_link(cursor, username, employee_id=None):
        if not username:
            return None
        account = load_user(cursor, username)
        if not account:
            raise HTTPException(400, 'Le compte à associer est introuvable.')
        canonical = account['username']
        cursor.execute('''SELECT employee_id FROM personnel_register
            WHERE linked_username=%s AND employee_id<>%s''', (canonical, employee_id or 0))
        if cursor.fetchone():
            raise HTTPException(409, 'Ce compte est déjà associé à une autre personne.')
        return canonical

    def displayed_details(cursor, details):
        """Add response-only account labels after audit snapshots have been written."""
        employees = [detail['employee'] for detail in details if isinstance(detail, dict)
                     and isinstance(detail.get('employee'), dict)]
        labels = load_display_names(cursor, [employee.get('linked_username') for employee in employees
                                            if employee.get('linked_username')])
        for employee in employees:
            linked = employee.get('linked_username')
            employee['linked_display_name'] = labels.get(linked, 'Ancien compte') if linked else None
        return details

    def transact(actor, change, *, require_super=False):
        db = get_db_connection()
        try:
            with db.cursor() as cursor:
                current = checked_admin(cursor, actor, lock=True)
                if require_super and not is_super_admin(current):
                    raise HTTPException(403, 'La suppression définitive est réservée au SuperAdmin.')
                result = change(cursor, current['username'])
                displayed_details(cursor, [result])
            db.commit()
            return result
        except IntegrityError:
            db.rollback()
            # The unique constraint also protects bindings against races and other callers.
            raise HTTPException(409, 'Conflit d’enregistrement. Vérifiez le compte associé puis actualisez.') from None
        except NirProtectionError:
            db.rollback()
            raise HTTPException(503, 'La protection des données sensibles est temporairement indisponible.') from None
        except OperationalError as exc:
            db.rollback()
            if exc.args and exc.args[0] in {1205, 1213}:
                raise HTTPException(409, 'Une autre modification est en cours. Actualisez puis réessayez.') from None
            raise
        except Exception:
            db.rollback()
            raise
        finally:
            db.close()

    @app.on_event('startup')
    def migrate_personnel_register():
        db = get_db_connection()
        try:
            with db.cursor() as cursor:
                ensure_personnel_schema(cursor)
            db.commit()
        finally:
            db.close()

    def account_personnel(cursor, target, can_manage_link):
        # This projection deliberately never selects NIR material or its presence
        # flag. Self-service grants no access to the register or its audit log.
        columns = ('employee_id', *EMPLOYEE_FIELDS, 'is_archived', 'version')
        cursor.execute(f"SELECT {','.join(columns)} FROM personnel_register WHERE linked_username=%s",
                       (target['username'],))
        employee = cursor.fetchone()
        periods = []
        if employee:
            employee = dict(employee)
            employee['is_archived'] = bool(employee['is_archived'])
            employee['registration_number'] = employee['employee_id']
            cursor.execute(f'SELECT {selected_columns(PERIOD_COLUMNS)} FROM personnel_register_periods '
                           'WHERE employee_id=%s ORDER BY entry_date DESC,period_id DESC',
                           (employee['employee_id'],))
            periods = [period_record(row) for row in cursor.fetchall()]
        return {'employee': employee, 'periods': periods, 'can_manage_link': can_manage_link,
                'account_display_name': employee['full_name'] if employee else person_display_name({'email': target.get('email')}),
                'account_email': target.get('email') or ''}

    @router.get('/account-personnel/{username}')
    def get_account_personnel(username: str, actor=Depends(get_current_user)):
        db = get_db_connection()
        try:
            with db.cursor() as cursor:
                current = checked_resource_actor(cursor, actor, load_user)
                target = load_user(cursor, username)
                if not target or (not is_admin(current) and current['username'] != target['username']):
                    raise HTTPException(403, 'Ces informations sont réservées à la personne et aux administrateurs.')
                return account_personnel(cursor, target, is_admin(current))
        finally:
            db.close()

    @router.get('/account-personnel/{username}/choices')
    def account_personnel_choices(username: str, q: str = Query(default='', max_length=200),
                                  limit: int = Query(default=30, ge=1, le=100), actor=Depends(get_current_user)):
        db = get_db_connection()
        try:
            with db.cursor() as cursor:
                checked_admin(cursor, actor)
                target = load_user(cursor, username)
                if not target:
                    raise HTTPException(404, 'Compte introuvable.')
                cursor.execute('''SELECT employee_id,full_name,job_title,personnel_type,version
                    FROM personnel_register WHERE is_archived=0
                    AND (linked_username IS NULL OR linked_username=%s)
                    AND (full_name LIKE %s ESCAPE '=' OR CAST(employee_id AS CHAR)=%s)
                    ORDER BY full_name,employee_id LIMIT %s''',
                    (target['username'], like_search(q.strip()), q.strip(), limit + 1))
                rows = cursor.fetchall()
                return {'employees': rows[:limit], 'has_more': len(rows) > limit}
        finally:
            db.close()

    @router.put('/account-personnel/{username}/link')
    def link_account_personnel(username: str, payload: AccountPersonnelLinkInput, actor=Depends(get_current_user)):
        def change(cursor, actor_username):
            target = load_user(cursor, username)
            if not target:
                raise HTTPException(404, 'Compte introuvable.')
            cursor.execute('SELECT employee_id FROM personnel_register WHERE linked_username=%s',
                           (target['username'],))
            linked = cursor.fetchone()
            current_id = linked['employee_id'] if linked else None
            if current_id != payload.expected_employee_id:
                raise HTTPException(409, 'L’association a changé. Actualisez avant de réessayer.')
            if (current_id is None) != (payload.expected_version is None):
                raise HTTPException(400, 'La version de l’association actuelle est requise.')
            if (payload.employee_id is None) != (payload.target_version is None):
                raise HTTPException(400, 'La version de la fiche à associer est requise.')
            # Users are already locked in a stable order by transact. Lock the
            # register records in ID order too; all changes and audits commit once.
            details = {key: detail_or_404(cursor, key, lock=True)
                       for key in sorted({key for key in (current_id, payload.employee_id) if key is not None})}
            if current_id is not None and details[current_id]['employee']['version'] != payload.expected_version:
                raise HTTPException(409, 'La fiche associée a changé. Actualisez avant de réessayer.')
            if payload.employee_id is not None:
                selected = details[payload.employee_id]['employee']
                if selected['version'] != payload.target_version:
                    raise HTTPException(409, 'La fiche choisie a changé. Actualisez avant de réessayer.')
                if selected['linked_username'] not in (None, target['username']):
                    raise HTTPException(409, 'Cette fiche est déjà associée à un autre compte.')
                if selected['is_archived'] and payload.employee_id != current_id:
                    raise HTTPException(409, 'Restaurez la fiche archivée avant de l’associer.')
            if current_id != payload.employee_id:
                if current_id is not None:
                    cursor.execute('UPDATE personnel_register SET linked_username=NULL,version=version+1 WHERE employee_id=%s',
                                   (current_id,))
                    after = detail_or_404(cursor, current_id)
                    append_personnel_event(cursor, current_id, actor_username, 'account_unlinked', details[current_id], after)
                if payload.employee_id is not None:
                    cursor.execute('UPDATE personnel_register SET linked_username=%s,version=version+1 WHERE employee_id=%s',
                                   (target['username'], payload.employee_id))
                    after = detail_or_404(cursor, payload.employee_id)
                    append_personnel_event(cursor, payload.employee_id, actor_username, 'account_linked',
                                           details[payload.employee_id], after)
            return account_personnel(cursor, target, True)
        return transact(actor, change)

    # Static routes must precede the employee-id route.
    @router.get('/personnel-register/capabilities')
    def capabilities(actor=Depends(get_current_user)):
        db = get_db_connection()
        try:
            with db.cursor() as cursor:
                checked_admin(cursor, actor)
                return {'nir_enabled': nir_enabled()}
        finally:
            db.close()

    @router.get('/personnel-register/accounts')
    def accounts(q: str = Query(default='', max_length=200),
                 limit: int = Query(default=30, ge=1, le=100), actor=Depends(get_current_user)):
        db = get_db_connection()
        try:
            with db.cursor() as cursor:
                checked_admin(cursor, actor)
                term = like_search(q.strip())
                cursor.execute('''SELECT u.username,u.email,u.is_active,u.login_deleted_at,
                    r.employee_id AS linked_employee_id,r.full_name FROM users u
                    LEFT JOIN personnel_register r ON r.linked_username=u.username
                    WHERE u.username LIKE %s ESCAPE '=' OR u.email LIKE %s ESCAPE '='
                        OR r.full_name LIKE %s ESCAPE '='
                    ORDER BY COALESCE(NULLIF(r.full_name,''),u.email),u.username LIMIT %s''',
                               (term, term, term, limit + 1))
                rows = cursor.fetchall()
                return {'accounts': [{**row, 'is_active': bool(row['is_active']),
                                      'display_name': person_display_name(row)} for row in rows[:limit]],
                        'has_more': len(rows) > limit, 'limit': limit}
        finally:
            db.close()

    @router.get('/personnel-register')
    def list_employees(q: str = Query(default='', max_length=200), personnel_type: PersonnelType = '',
                       archive: ArchiveFilter = 'active', status: EmploymentStatus = 'all',
                       page: int = Query(default=1, ge=1, le=1_000_000),
                       page_size: int = Query(default=25, ge=1, le=100), actor=Depends(get_current_user)):
        db = get_db_connection()
        try:
            with db.cursor() as cursor:
                checked_admin(cursor, actor)
                clauses, params = [], []
                if q.strip():
                    clauses.append("(e.full_name LIKE %s ESCAPE '=' OR e.linked_username LIKE %s ESCAPE '=' "
                                   'OR CAST(e.employee_id AS CHAR)=%s)')
                    params.extend((like_search(q.strip()), like_search(q.strip()), q.strip()))
                if personnel_type:
                    clauses.append('e.personnel_type=%s')
                    params.append(personnel_type)
                if archive != 'all':
                    clauses.append('e.is_archived=%s')
                    params.append(archive == 'archived')
                if status != 'all':
                    clauses.append(f'({_STATUS_FILTER})=%s')
                    params.extend((paris_today(), paris_today(), status))
                where = ' WHERE ' + ' AND '.join(clauses) if clauses else ''
                base = f'FROM personnel_register e {_LATEST_PERIOD_JOIN}{where}'
                cursor.execute(f'SELECT COUNT(*) AS total {base}', tuple(params))
                total = cursor.fetchone()['total']
                employee_columns = selected_columns(EMPLOYEE_COLUMNS, table='e')
                period_columns = selected_columns(PERIOD_COLUMNS, table='p', prefix='period_')
                cursor.execute(
                    f'SELECT {employee_columns},{period_columns} {base} ORDER BY e.employee_id LIMIT %s OFFSET %s',
                    (*params, page_size, (page - 1) * page_size),
                )
                employees = []
                for row in cursor.fetchall():
                    period = (period_record({key: row[f'period_{key}'] for key in PERIOD_COLUMNS})
                              if row['period_period_id'] is not None else None)
                    employees.append(employee_record(row, period))
                displayed_details(cursor, [{'employee': employee} for employee in employees])
                return {'employees': employees, 'total': total, 'page': page, 'page_size': page_size}
        finally:
            db.close()

    @router.post('/personnel-register', status_code=201)
    def create_employee(payload: EmployeeCreate, actor=Depends(get_current_user)):
        def change(cursor, actor_username):
            nir_requested = 'social_security_number' in payload.model_fields_set
            if nir_requested:
                require_nir_storage()
            values = payload.model_dump(include=set(EMPLOYEE_FIELDS))
            values['linked_username'] = canonical_link(cursor, values['linked_username'])
            placeholders = ','.join(['%s'] * len(EMPLOYEE_FIELDS))
            cursor.execute(f"INSERT INTO personnel_register ({','.join(EMPLOYEE_FIELDS)}) VALUES ({placeholders})",
                           tuple(values[key] for key in EMPLOYEE_FIELDS))
            employee_id = cursor.lastrowid
            if nir_requested and payload.social_security_number is not None:
                envelope = encrypt_social_security_number(payload.social_security_number.get_secret_value(), employee_id)
                cursor.execute('UPDATE personnel_register SET social_security_number_encrypted=%s WHERE employee_id=%s',
                               (envelope, employee_id))
            if payload.initial_period is not None:
                insert_period(cursor, employee_id, payload.initial_period.model_dump())
            after = detail_or_404(cursor, employee_id)
            append_personnel_event(cursor, employee_id, actor_username, 'registered', None, after)
            if after['employee']['has_social_security_number']:
                append_personnel_event(cursor, employee_id, actor_username, 'social_security_number_changed',
                                       sensitive_metadata(employee_id, False), sensitive_metadata(employee_id, True))
            return after
        return transact(actor, change)

    @router.get('/personnel-register/{employee_id}')
    def get_employee(employee_id: int = Path(ge=1, le=MAX_IDENTIFIER), actor=Depends(get_current_user)):
        db = get_db_connection()
        try:
            with db.cursor() as cursor:
                checked_admin(cursor, actor)
                detail = detail_or_404(cursor, employee_id)
                displayed_details(cursor, [detail])
                return detail
        finally:
            db.close()

    @router.put('/personnel-register/{employee_id}')
    def update_employee(payload: EmployeeUpdate, employee_id: int = Path(ge=1, le=MAX_IDENTIFIER),
                        actor=Depends(get_current_user)):
        def change(cursor, actor_username):
            before = editable_detail(cursor, employee_id, payload.expected_version)
            values = payload.model_dump(include=set(EMPLOYEE_FIELDS))
            values['linked_username'] = canonical_link(cursor, values['linked_username'], employee_id)
            details_changed = any(before['employee'][key] != values[key] for key in EMPLOYEE_FIELDS)
            nir_requested = 'social_security_number' in payload.model_fields_set
            if nir_requested:
                require_nir_storage()
            nir_changed = nir_requested and (payload.social_security_number is not None
                                           or before['employee']['has_social_security_number'])
            if details_changed or nir_changed:
                assignments = ','.join(f'{key}=%s' for key in EMPLOYEE_FIELDS)
                parameters = [values[key] for key in EMPLOYEE_FIELDS]
                if nir_changed:
                    envelope = (encrypt_social_security_number(payload.social_security_number.get_secret_value(), employee_id)
                                if payload.social_security_number is not None else None)
                    assignments += ',social_security_number_encrypted=%s'
                    parameters.append(envelope)
                cursor.execute(f'UPDATE personnel_register SET {assignments},version=version+1 WHERE employee_id=%s',
                               (*parameters, employee_id))
                after = detail_or_404(cursor, employee_id)
                if details_changed:
                    append_personnel_event(cursor, employee_id, actor_username, 'details_updated', before, after)
                if nir_changed:
                    append_personnel_event(cursor, employee_id, actor_username, 'social_security_number_changed',
                                           sensitive_metadata(employee_id, before['employee']['has_social_security_number']),
                                           sensitive_metadata(employee_id, after['employee']['has_social_security_number']))
                return after
            return before
        return transact(actor, change)

    @router.post('/personnel-register/{employee_id}/social-security-number/reveal')
    def reveal_social_security_number(employee_id: int = Path(ge=1, le=MAX_IDENTIFIER),
                                      actor=Depends(get_current_user)):
        def change(cursor, actor_username):
            require_nir_storage()
            detail = detail_or_404(cursor, employee_id, lock=True)
            cursor.execute('SELECT social_security_number_encrypted FROM personnel_register WHERE employee_id=%s',
                           (employee_id,))
            envelope = cursor.fetchone()['social_security_number_encrypted']
            if not detail['employee']['has_social_security_number'] or envelope is None:
                raise HTTPException(404, 'Aucun numéro enregistré pour cette personne.')
            number = decrypt_social_security_number(envelope, employee_id)
            metadata = sensitive_metadata(employee_id, True)
            append_personnel_event(cursor, employee_id, actor_username, 'social_security_number_viewed', metadata, metadata)
            return {'social_security_number': number}
        # transact commits the successful read audit before returning its short-lived plaintext.
        return transact(actor, change)

    @router.post('/personnel-register/{employee_id}/periods', status_code=201)
    def add_period(payload: PeriodMutation, employee_id: int = Path(ge=1, le=MAX_IDENTIFIER),
                   actor=Depends(get_current_user)):
        def change(cursor, actor_username):
            before = editable_detail(cursor, employee_id, payload.expected_version)
            insert_period(cursor, employee_id, payload.model_dump(include=set(PERIOD_FIELDS)))
            touch_employee(cursor, employee_id)
            after = detail_or_404(cursor, employee_id)
            append_personnel_event(cursor, employee_id, actor_username, 'period_added', before, after)
            return after
        return transact(actor, change)

    @router.put('/personnel-register/{employee_id}/periods/{period_id}')
    def update_period(payload: PeriodMutation, employee_id: int = Path(ge=1, le=MAX_IDENTIFIER),
                      period_id: int = Path(ge=1, le=MAX_IDENTIFIER), actor=Depends(get_current_user)):
        def change(cursor, actor_username):
            before = editable_detail(cursor, employee_id, payload.expected_version)
            period = next((item for item in before['periods'] if item['period_id'] == period_id), None)
            if not period:
                raise HTTPException(404, 'Période d’emploi introuvable pour cette personne.')
            values = payload.model_dump(include=set(PERIOD_FIELDS))
            if any(period[key] != values[key] for key in PERIOD_FIELDS):
                assignments = ','.join(f'{key}=%s' for key in PERIOD_FIELDS)
                cursor.execute(
                    f'UPDATE personnel_register_periods SET {assignments} WHERE employee_id=%s AND period_id=%s',
                    (*(values[key] for key in PERIOD_FIELDS), employee_id, period_id),
                )
                touch_employee(cursor, employee_id)
                after = detail_or_404(cursor, employee_id)
                append_personnel_event(cursor, employee_id, actor_username, 'period_updated', before, after)
                return after
            return before
        return transact(actor, change)

    @router.patch('/personnel-register/{employee_id}/archive')
    def archive_employee(payload: ArchiveMutation, employee_id: int = Path(ge=1, le=MAX_IDENTIFIER),
                         actor=Depends(get_current_user)):
        def change(cursor, actor_username):
            before = editable_detail(cursor, employee_id, payload.expected_version)
            if before['employee']['is_archived'] != payload.is_archived:
                cursor.execute('UPDATE personnel_register SET is_archived=%s,version=version+1 WHERE employee_id=%s',
                               (payload.is_archived, employee_id))
                after = detail_or_404(cursor, employee_id)
                action = 'archived' if payload.is_archived else 'restored'
                append_personnel_event(cursor, employee_id, actor_username, action, before, after)
                return after
            return before
        return transact(actor, change)

    @router.get('/personnel-register/{employee_id}/history')
    def history(employee_id: int = Path(ge=1, le=MAX_IDENTIFIER),
                page: int = Query(default=1, ge=1, le=1_000_000),
                page_size: int = Query(default=20, ge=1, le=100), actor=Depends(get_current_user)):
        db = get_db_connection()
        try:
            with db.cursor() as cursor:
                checked_admin(cursor, actor)
                cursor.execute('SELECT employee_id FROM personnel_register WHERE employee_id=%s', (employee_id,))
                if not cursor.fetchone():
                    raise HTTPException(404, 'Salarié introuvable.')
                cursor.execute('SELECT COUNT(*) AS total FROM personnel_register_events WHERE employee_id=%s',
                               (employee_id,))
                total = cursor.fetchone()['total']
                cursor.execute('''SELECT event_id,actor_username,action,before_snapshot,after_snapshot,
                    UNIX_TIMESTAMP(created_at) AS created_at
                    FROM personnel_register_events WHERE employee_id=%s
                    ORDER BY event_id DESC LIMIT %s OFFSET %s''', (employee_id, page_size, (page - 1) * page_size))
                events = []
                for row in cursor.fetchall():
                    events.append({
                        'event_id': row['event_id'], 'actor_username': row['actor_username'],
                        'action': row['action'], 'created_at': utc_timestamp(row['created_at']),
                        'before': json.loads(row['before_snapshot']) if row['before_snapshot'] else None,
                        'after': json.loads(row['after_snapshot']),
                    })
                related = {event['actor_username'] for event in events if event.get('actor_username')}
                for event in events:
                    for snapshot in (event['before'], event['after']):
                        employee = snapshot.get('employee', {}) if isinstance(snapshot, dict) else {}
                        if employee.get('linked_username'):
                            related.add(employee['linked_username'])
                labels = load_display_names(cursor, related)
                for event in events:
                    event['actor_display_name'] = labels.get(event['actor_username'], 'Ancien compte')
                    for snapshot in (event['before'], event['after']):
                        employee = snapshot.get('employee', {}) if isinstance(snapshot, dict) else {}
                        linked = employee.get('linked_username')
                        if linked:
                            # An explicit historical association carries its own recorded name.
                            employee['linked_display_name'] = (str(employee.get('full_name') or '').strip()
                                                               or labels.get(linked, 'Ancien compte'))
                        elif 'linked_username' in employee:
                            employee['linked_display_name'] = None
                return {'events': events, 'total': total, 'page': page, 'page_size': page_size}
        finally:
            db.close()

    @router.delete('/personnel-register/{employee_id}')
    def delete_employee(payload: PersonnelDeleteInput, employee_id: int = Path(ge=1, le=MAX_IDENTIFIER),
                        actor=Depends(get_current_user)):
        def change(cursor, actor_username):
            before = editable_detail(cursor, employee_id, payload.expected_version)
            employee = before['employee']
            if not employee['is_archived']:
                raise HTTPException(409, 'Archivez cette fiche avant sa suppression définitive.')
            if payload.confirmation not in {str(employee_id), employee['full_name']}:
                raise HTTPException(400, 'Confirmez le numéro ou le nom exact de cette fiche.')
            metadata = delete_personnel_records(cursor, employee_id)
            append_deletion_event(cursor, actor_username, 'personnel', employee_id, metadata)
            return {'status': 'success', 'deleted_employee_id': employee_id}
        return transact(actor, change, require_super=True)

    app.include_router(router)
