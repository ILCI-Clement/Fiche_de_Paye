"""Admin-only employee register with linked contract date synchronization.

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
from fastapi.responses import JSONResponse, Response
from pydantic import BaseModel, ConfigDict, Field
from pymysql.err import IntegrityError, OperationalError

try:
    from .personnel_contract_sync import detach_contract_links, sync_account_contracts, sync_register_period
    from .personnel_export import MAX_EXPORT_ROWS, generate_personnel_excel
    from .personnel_rules import (
        ArchiveFilter, ArchiveMutation, EmployeeCreate, EmployeeUpdate,
        EmploymentStatus, MAX_IDENTIFIER, PeriodInput, PeriodMutation, PersonnelType, like_search, paris_today, registration_missing, entry_precedes_birth,
    )
    from .personnel_store import (
        EMPLOYEE_COLUMNS, EMPLOYEE_FIELDS, PERIOD_COLUMNS, PERIOD_FIELDS,
        append_personnel_event, delete_personnel_records, employee_record, ensure_personnel_schema, insert_period,
        load_employee_detail, period_record, selected_columns,
        sensitive_metadata, touch_employee, utc_timestamp,
    )
    from .personnel_sensitive import (
        NirProtectionError, decrypt_social_security_number, encrypt_social_security_number, nir_storage_ready,
    )
    from .destructive_audit import append_deletion_event
    from .role_policy import checked_resource_actor, is_admin, is_super_admin
    from .person_names import load_display_names, load_person_accounts, person_display_name
except ImportError:
    from personnel_contract_sync import detach_contract_links, sync_account_contracts, sync_register_period
    from personnel_export import MAX_EXPORT_ROWS, generate_personnel_excel
    from personnel_rules import (
        ArchiveFilter, ArchiveMutation, EmployeeCreate, EmployeeUpdate,
        EmploymentStatus, MAX_IDENTIFIER, PeriodInput, PeriodMutation, PersonnelType, like_search, paris_today, registration_missing, entry_precedes_birth,
    )
    from personnel_store import (
        EMPLOYEE_COLUMNS, EMPLOYEE_FIELDS, PERIOD_COLUMNS, PERIOD_FIELDS,
        append_personnel_event, delete_personnel_records, employee_record, ensure_personnel_schema, insert_period,
        load_employee_detail, period_record, selected_columns,
        sensitive_metadata, touch_employee, utc_timestamp,
    )
    from personnel_sensitive import (
        NirProtectionError, decrypt_social_security_number, encrypt_social_security_number, nir_storage_ready,
    )
    from destructive_audit import append_deletion_event
    from role_policy import checked_resource_actor, is_admin, is_super_admin
    from person_names import load_display_names, load_person_accounts, person_display_name


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
    WHEN p.actual_exit_date IS NOT NULL AND p.actual_exit_date<%s THEN 'left'
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


class PersonnelExportInput(BaseModel):
    model_config = ConfigDict(extra='forbid', str_strip_whitespace=True, hide_input_in_errors=True)
    q: str = Field(default='', max_length=200)
    personnel_type: PersonnelType = ''
    archive: ArchiveFilter = 'active'
    status: EmploymentStatus = 'all'
    confirm_sensitive_export: bool = Field(strict=True)


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
            with db.cursor() as cursor:
                cursor.execute('SELECT username FROM users ORDER BY username FOR UPDATE')
                cursor.fetchall()
                cursor.execute('SELECT linked_username FROM personnel_register WHERE linked_username IS NOT NULL ORDER BY employee_id')
                for row in cursor.fetchall():
                    sync_account_contracts(cursor, row['linked_username'], 'system:contract-sync')
            db.commit()
        except Exception:
            db.rollback()
            raise
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
                'account_display_name': load_display_names(cursor, [target['username']]).get(
                    target['username'], person_display_name({'email': target.get('email')})),
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
                    detach_contract_links(cursor, current_id)
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
                    sync_account_contracts(cursor, target['username'], actor_username)
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
                directory = load_person_accounts(cursor, [row['username'] for row in rows[:limit]])
                return {'accounts': [{**row, 'is_active': bool(row['is_active']),
                                      'full_name': directory.get(row['username'], {}).get('full_name'),
                                      'display_name': directory.get(row['username'], {}).get('display_name', person_display_name(row))} for row in rows[:limit]],
                        'has_more': len(rows) > limit, 'limit': limit}
        finally:
            db.close()

    def employee_filter(q, personnel_type, archive, status):
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
            today = paris_today()
            params.extend((today, today, status))
        where = ' WHERE ' + ' AND '.join(clauses) if clauses else ''
        return where, params

    @router.post('/personnel-register/export.xlsx')
    def export_employees(payload: PersonnelExportInput, actor=Depends(get_current_user)):
        def generate(cursor, actor_username):
            if not payload.confirm_sensitive_export:
                raise HTTPException(400, 'Confirmez l’export des données personnelles et des numéros de sécurité sociale.')
            require_nir_storage()
            where, params = employee_filter(payload.q, payload.personnel_type, payload.archive, payload.status)
            columns = ','.join(f'e.{key}' for key in (
                'employee_id', 'full_name', 'nationality', 'sex', 'address', 'birth_date',
                'job_title', 'qualification', 'social_security_number_encrypted'))
            history_columns = ','.join(f'h.{key}' for key in PERIOD_FIELDS)
            cursor.execute(f'''SELECT {columns},{history_columns}
                FROM personnel_register e {_LATEST_PERIOD_JOIN}
                LEFT JOIN personnel_register_periods h ON h.employee_id=e.employee_id
                {where} ORDER BY e.employee_id,h.entry_date,h.period_id LIMIT %s''',
                           (*params, MAX_EXPORT_ROWS + 1))
            rows = cursor.fetchall()
            if not rows:
                raise HTTPException(404, 'Aucune personne ne correspond aux filtres à exporter.')
            if len(rows) > MAX_EXPORT_ROWS:
                raise HTTPException(400, f'L’export est limité à {MAX_EXPORT_ROWS} lignes. Précisez les filtres ; aucun fichier partiel n’a été généré.')
            numbers, counts = {}, {}
            for row in rows:
                identifier = row['employee_id']
                if identifier not in numbers:
                    envelope = row.get('social_security_number_encrypted')
                    numbers[identifier] = decrypt_social_security_number(envelope, identifier) if envelope is not None else ''
                row['social_security_number'] = numbers[identifier]
                row.pop('social_security_number_encrypted', None)
                counts[identifier] = counts.get(identifier, 0) + 1
            content = generate_personnel_excel(rows)
            # Commit one metadata-only event per person before releasing plaintext.
            # Never persist the generated file, filter text, number or envelope.
            for identifier, row_count in counts.items():
                metadata = sensitive_metadata(identifier, bool(numbers[identifier]))
                metadata['export'] = {'format': 'xlsx', 'row_count': row_count,
                                      'includes_social_security_number': True}
                append_personnel_event(cursor, identifier, actor_username, 'excel_exported', metadata, metadata)
            filename = f'registre-du-personnel-{paris_today().isoformat()}.xlsx'
            return Response(content, media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                            headers={'Content-Disposition': f'attachment; filename="{filename}"'})
        return transact(actor, generate)

    @router.get('/personnel-register')
    def list_employees(q: str = Query(default='', max_length=200), personnel_type: PersonnelType = '',
                       archive: ArchiveFilter = 'active', status: EmploymentStatus = 'all',
                       page: int = Query(default=1, ge=1, le=1_000_000),
                       page_size: int = Query(default=25, ge=1, le=100), actor=Depends(get_current_user)):
        db = get_db_connection()
        try:
            with db.cursor() as cursor:
                checked_admin(cursor, actor)
                where, params = employee_filter(q, personnel_type, archive, status)
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
            cursor.execute(f"INSERT INTO personnel_register ({','.join(EMPLOYEE_FIELDS)},registration_schema_version) VALUES ({placeholders},2)",
                           tuple(values[key] for key in EMPLOYEE_FIELDS))
            employee_id = cursor.lastrowid
            if nir_requested and payload.social_security_number is not None:
                envelope = encrypt_social_security_number(payload.social_security_number.get_secret_value(), employee_id)
                cursor.execute('UPDATE personnel_register SET social_security_number_encrypted=%s WHERE employee_id=%s',
                               (envelope, employee_id))
            if payload.initial_period is not None:
                period_id = insert_period(cursor, employee_id, payload.initial_period.model_dump())
            if values['linked_username']:
                sync_account_contracts(cursor, values['linked_username'], actor_username,
                                       match_legacy_period=False, initial_period_id=period_id)
                if payload.initial_period is not None:
                    # First association uses account dates, including an existing period match.
                    cursor.execute('SELECT account_contract_id FROM personnel_register_periods WHERE period_id=%s', (period_id,))
                    if cursor.fetchone()['account_contract_id'] is None:
                        sync_register_period(cursor, employee_id, period_id)
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
            # An older client cannot clear structured names it never received.
            for key in ('first_name', 'last_name'):
                if key not in payload.model_fields_set:
                    values[key] = before['employee'][key]
            if values['first_name'] and values['last_name']:
                values['full_name'] = f"{values['first_name']} {values['last_name']}"
                if len(values['full_name']) > 200:
                    raise HTTPException(400, 'Le nom et les prénoms sont limités à 200 caractères au total.')
            values['linked_username'] = canonical_link(cursor, values['linked_username'], employee_id)
            if before['employee']['registration_schema_version'] >= 2:
                missing = registration_missing(values, before['employee']['latest_period'])
                if missing:
                    raise HTTPException(400, 'Champs obligatoires : ' + ', '.join(missing) + '.')
            if values['birth_date'] != before['employee']['birth_date'] and any(
                    entry_precedes_birth(p['entry_date'], values['birth_date']) for p in before['periods']):
                raise HTTPException(400, 'La date d’entrée ne peut pas précéder la naissance.')
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
                if before['employee']['linked_username'] != values['linked_username']:
                    detach_contract_links(cursor, employee_id)
                    if values['linked_username']:
                        sync_account_contracts(cursor, values['linked_username'], actor_username)
                    after = detail_or_404(cursor, employee_id)
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
            if payload.entry_date is None:
                raise HTTPException(400, 'Renseignez la date d’entrée de la nouvelle période.')
            if entry_precedes_birth(payload.entry_date, before['employee']['birth_date']):
                raise HTTPException(400, 'La date d’entrée ne peut pas précéder la naissance.')
            period_id = insert_period(cursor, employee_id, payload.model_dump(include=set(PERIOD_FIELDS)))
            sync_register_period(cursor, employee_id, period_id)
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
            try:
                values = PeriodInput(**{key: values[key] if key in payload.model_fields_set else period[key]
                                        for key in PERIOD_FIELDS}).storage_values()
            except ValueError:
                raise HTTPException(400, 'Vérifiez les dates et les champs obligatoires des situations particulières.') from None
            if before['employee']['registration_schema_version'] >= 2 and values['entry_date'] is None:
                raise HTTPException(400, 'La date d’entrée est obligatoire.')
            if values['entry_date'] != period['entry_date'] and entry_precedes_birth(values['entry_date'], before['employee']['birth_date']):
                raise HTTPException(400, 'La date d’entrée ne peut pas précéder la naissance.')
            if any(period[key] != values[key] for key in PERIOD_FIELDS):
                assignments = ','.join(f'{key}=%s' for key in PERIOD_FIELDS)
                cursor.execute(
                    f'UPDATE personnel_register_periods SET {assignments} WHERE employee_id=%s AND period_id=%s',
                    (*(values[key] for key in PERIOD_FIELDS), employee_id, period_id),
                )
                if any(period[key] != values[key] for key in ('entry_date', 'exit_date', 'contract_type')):
                    sync_register_period(cursor, employee_id, period_id)
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
                    event['actor_display_name'] = ('Synchronisation automatique'
                        if event['actor_username'] == 'system:contract-sync'
                        else labels.get(event['actor_username'], 'Ancien compte'))
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
