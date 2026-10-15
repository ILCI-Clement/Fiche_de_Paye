"""Loopback test backend using the private production snapshot, never its DB/server keys."""
import json
import os
from pathlib import Path

from fastapi import HTTPException
from .local_signing_policy import LocalSigningPolicy

ROOT = Path(__file__).resolve().parents[1]
STATE = ROOT / '.local-test' / 'rebuild'
runtime = json.loads((STATE / 'runtime.json').read_text(encoding='utf-8-sig'))
if runtime['database'] != 'presence_local_rebuild':
    raise RuntimeError('Unexpected local test database.')

os.environ.update({
    'PRESENCE_API_TOKEN': runtime['api_token'], 'PRESENCE_DB_HOST': '127.0.0.1',
    'PRESENCE_DB_USER': 'root', 'PRESENCE_DB_PASSWORD': runtime['db_password'],
    'PRESENCE_DB_DATABASE': runtime['database'],
    'PRESENCE_SMTP_SERVER': '127.0.0.1', 'PRESENCE_SMTP_PORT': '1',
    'PRESENCE_SMTP_USER': 'disabled', 'PRESENCE_SMTP_PASSWORD': 'disabled',
    'PRESENCE_APP_URL': 'http://127.0.0.1:5173',
    'PRESENCE_TRANSPORT_RECEIPTS_DIR': str(STATE / 'snapshot' / 'documents'),
    'PRESENCE_ESIGN_DOCUMENTS_DIR': str(STATE / 'esign-documents'),
})
# No inherited signing credential: only the private, explicitly selected local policy.
os.environ.pop('CLAWSHOW_ESIGN_API_KEY', None)
os.environ.pop('PRESENCE_ESIGN_PUBLIC_URL', None)
signing_policy = LocalSigningPolicy(runtime)
selected_key = signing_policy.key()
if selected_key:
    os.environ['CLAWSHOW_ESIGN_API_KEY'] = selected_key

soffice = STATE / 'libreoffice' / 'program' / 'soffice.exe'
if soffice.exists():
    os.environ['PATH'] = str(soffice.parent) + os.pathsep + os.environ['PATH']

from . import main
from .attendance_api import register_attendance_routes
from .attendance_signatures import register_signature_preflight
from .signature_attempts import SignatureAttempts
from urllib.parse import urlsplit

main.DB_CONFIG.update(host='127.0.0.1', port=3307, database=runtime['database'])
main.start_contract_reminder_worker = lambda: None
main.send_password_reset_email = lambda *_args: False
main.send_contract_end_reminder_email = lambda *_args: False
app = main.app
app.state.local_snapshot = True
app.state.esign_recipient_policy = signing_policy.check


@app.middleware('http')
async def local_only(request, call_next):
    if request.client and request.client.host not in {'127.0.0.1', '::1', 'testclient'}:
        from fastapi.responses import JSONResponse
        return JSONResponse({'detail': 'Local test environment only.'}, status_code=403)
    if request.url.path in {'/forgot-password', '/contract-reminders/send', '/contract-reminders/resend'}:
        from fastapi.responses import JSONResponse
        return JSONResponse({'detail': 'Les e-mails sont désactivés dans cette copie locale.'}, status_code=403)
    if request.url.path == '/create-user':
        from fastapi.responses import JSONResponse
        return JSONResponse({'detail': 'Aucun compte fictif dans cette copie de test.'}, status_code=403)
    return await call_next(request)


register_attendance_routes(app, get_db_connection=main.get_db_connection,
    get_current_user=main.get_current_user, load_user=main.load_user,
    can_manage=main.user_can_manage, save_config=main.save_config, save_checked=main.save_config_checked)

def signing_readiness():
    mode = signing_policy.mode
    try:
        signing_policy.check()
    except HTTPException as error:
        return {'mode': mode, 'can_send': False, 'reason': str(error.detail)}
    key = os.getenv('CLAWSHOW_ESIGN_API_KEY', '')
    public = str(runtime.get('esign_public_url') or '').rstrip('/')
    parsed = urlsplit(public)
    valid_url = parsed.scheme == 'https' and bool(parsed.hostname) and parsed.hostname not in {'localhost', '127.0.0.1', '::1'} and not parsed.query and not parsed.fragment and not parsed.username
    expected_prefix = 'sk_live_' if signing_policy.live else 'sk_test_'
    if not key.startswith(expected_prefix):
        return {'mode': mode, 'can_send': False, 'reason': 'La clé du mode de signature sélectionné n’est pas configurée.'}
    if not valid_url:
        return {'mode': mode, 'can_send': False, 'reason': 'Configurez une URL HTTPS publique dédiée aux PDF de ce backend de test.'}
    os.environ['PRESENCE_ESIGN_PUBLIC_URL'] = public
    if signing_policy.live:
        return {'mode': mode, 'can_send': True, 'reason': 'Test Production : de vrais e-mails seront envoyés aux deux signataires autorisés. Une demande de signature réelle sera créée et consommera le quota.'}
    return {'mode': mode, 'can_send': True, 'reason': 'Sandbox uniquement : aucun e-mail envoyé, documents SPECIMEN sans valeur juridique.'}


register_signature_preflight(app, get_current_user=main.get_current_user,
    get_db_connection=main.get_db_connection, load_user=main.load_user,
    require_target_access=main.require_target_access, get_config=main.get_config,
    readiness=signing_readiness, sender=main.send_fiche_for_signature,
    attempts=SignatureAttempts(STATE / 'signature-attempts.sqlite3'), recipient_policy=signing_policy.check)


@app.get('/local-environment')
def environment():
    return {'mode': 'local-snapshot', 'mail_enabled': False, 'signing_mode': signing_policy.mode,
            'signature_email_test': signing_policy.live,
            'snapshot_date': '2026-10-14'}
