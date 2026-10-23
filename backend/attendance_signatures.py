"""React signing adapter: saved sheets, existing dual-sign service, sandbox guard."""
import copy
import base64
import hashlib
import json
import time
from fastapi import Depends, HTTPException
from pydantic import BaseModel, Field
from contract_periods import period_sheets
from fiche_exports import validate_fiche, generate_fiche_file
from fiche_serialization import deserialize_dates
from .signature_flow import resolve_manager
from .esign import ESignFicheRequest
from .person_names import person_display_name


class SignatureSelection(BaseModel):
    ids: list[str] = Field(min_length=1, max_length=100)
    month: int = Field(ge=1, le=12)
    year: int = Field(ge=2000, le=2100)


class SignatureSubmit(BaseModel):
    id: str
    month: int = Field(ge=1, le=12)
    year: int = Field(ge=2000, le=2100)
    fingerprint: str = Field(min_length=64, max_length=64)
    request_id: str = Field(min_length=16, max_length=128, pattern=r'^[a-zA-Z0-9_-]+$')


def register_signature_preflight(app, *, get_current_user, get_db_connection, load_user,
                                 require_target_access, get_config, readiness, sender, attempts, recipient_policy=None):
    def inspect(employee, month, year, actor):
        restored = deserialize_dates(copy.deepcopy(employee))
        validate_fiche(restored, require_email=True,
                       require_first_name=employee.get('type') == 'Stagiaire', check_exceptions=True)
        email = str(employee.get('email_employe') or '')
        manager = resolve_manager(email, actor, get_db_connection, load_user, require_target_access)
        if recipient_policy:
            recipient_policy(email, manager['email'])
        name = manager.get('employee_name') or person_display_name({'email': email})
        digest = hashlib.sha256(json.dumps({'employee': employee, 'manager': manager, 'mode': readiness().get('mode', 'sandbox-only'),
            'month': month, 'year': year}, sort_keys=True, default=str).encode()).hexdigest()
        return {'name': name, 'employee_email': email, 'manager_name': manager['name'],
                'manager_email': manager['email'], 'fingerprint': digest}

    def records_for(actor, month, year):
        if not {'Admin', 'Responsable'} & set(actor['role_tags']):
            raise HTTPException(403, 'Accès réservé aux administrateurs et responsables.')
        config = get_config(actor['username'], actor)
        return {str(e['id']): e for e in period_sheets(config.get('employes_data', []), month, year)
                if not e.get('is_archived')}

    @app.post('/attendance/signature-preview')
    def preview(payload: SignatureSelection, actor=Depends(get_current_user)):
        records = records_for(actor, payload.month, payload.year)
        results = []
        for sheet_id in dict.fromkeys(payload.ids):
            employee = records.get(sheet_id)
            if not employee:
                raise HTTPException(400, 'Une fiche sélectionnée est absente du mois affiché.')
            name = person_display_name({'email': employee.get('email_employe')})
            result = {'id': sheet_id, 'name': name, 'employee_email': str(employee.get('email_employe') or ''), 'manager_name': '', 'manager_email': '', 'error': ''}
            try:
                result.update(inspect(employee, payload.month, payload.year, actor))
            except (HTTPException, ValueError, TypeError, KeyError) as error:
                result['error'] = str(error.detail) if isinstance(error, HTTPException) else str(error)
            results.append(result)
        archive = getattr(app.state, 'signed_fiches', None)
        if archive:
            from .signed_fiches import callback_url
        return {'items': results, 'mode': 'sandbox-only', **readiness(),
                'automatic_archive': bool(archive and callback_url())}

    @app.post('/attendance/signature-send')
    def send(payload: SignatureSubmit, actor=Depends(get_current_user)):
        records = records_for(actor, payload.month, payload.year)
        ready = readiness()
        if not ready['can_send']:
            raise HTTPException(503, ready['reason'])
        employee = records.get(payload.id)
        if not employee:
            raise HTTPException(400, 'La fiche ne fait pas partie du mois affiché.')
        try:
            details = inspect(employee, payload.month, payload.year, actor)
            if details['fingerprint'] != payload.fingerprint:
                raise HTTPException(409, 'La fiche ou son responsable a changé. Actualisez la préparation.')
            db = get_db_connection()
            try:
                with db.cursor() as cursor:
                    cursor.execute('SELECT username FROM users WHERE LOWER(email)=LOWER(%s)', (details['employee_email'],))
                    subjects = cursor.fetchall()
                subject_username = subjects[0]['username'] if len(subjects) == 1 else None
            finally:
                db.close()
            existing = attempts.begin(actor['username'], payload.request_id, payload.fingerprint)
            if existing is not None:
                return existing
            try:
                started = time.monotonic()
                item = generate_fiche_file(deserialize_dates(copy.deepcopy(employee)), payload.month, payload.year)
                generated_at = time.monotonic()
                result = sender(ESignFicheRequest(recipient_email=details['employee_email'], employee_name=details['name'],
                    month=payload.month, year=payload.year, filename=item['filename'],
                    file_b64=base64.b64encode(item['data']).decode()), actor)
                response = {'id': payload.id, 'status': 'accepted', 'document_id': result.get('document_id', ''),
                            'archive_id': result.get('archive_id'),
                            'profile_subject': {'username': subject_username, 'month': payload.month, 'year': payload.year},
                            'mode': ready.get('mode', 'sandbox-only'),
                            'message': result.get('message', ''), 'request_id': payload.request_id,
                            'signing_links': result.get('signing_links', []),
                            'signing_actions': result.get('signing_actions', []),
                            'timings': {'generation_seconds': round(generated_at - started, 2),
                                        'pdf_and_provider_seconds': round(time.monotonic() - generated_at, 2),
                                        'total_seconds': round(time.monotonic() - started, 2),
                                        **result.get('timings', {})}}
                attempts.finish(actor['username'], payload.request_id, 'accepted', response)
                return response
            except Exception:
                # Even an error may follow a provider acceptance; do not retry automatically.
                attempts.finish(actor['username'], payload.request_id, 'uncertain')
                raise
        except (ValueError, TypeError, KeyError) as error:
            raise HTTPException(422, str(error)) from None
