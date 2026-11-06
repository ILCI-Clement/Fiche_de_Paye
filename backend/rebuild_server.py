"""Staged React server entry. Do not start against production before release approval.

Uses the original server environment and data paths, never local snapshot/runtime.json.
Run backend.rebuild_server:app from the complete repository root.
"""
import os
from pathlib import Path

from . import main
from .attendance_api import register_attendance_routes
from .attendance_signatures import register_signature_preflight
from .signature_attempts import SignatureAttempts
from .server_signing_readiness import server_signing_readiness
from .person_profiles import register_profile_routes
from .personnel_register import register_personnel_routes
from .resource_maintenance import register_resource_maintenance

app = main.app
from .system_storage import register_storage_routes, application_storage_paths
register_storage_routes(app, get_db_connection=main.get_db_connection,
                        get_current_user=main.get_current_user, load_user=main.load_user,
                        paths=application_storage_paths(main))
from .personnel_snapshots import register_snapshot_routes
register_snapshot_routes(app, get_db_connection=main.get_db_connection,
                         get_current_user=main.get_current_user, load_user=main.load_user)
register_personnel_routes(app, get_db_connection=main.get_db_connection,
                          get_current_user=main.get_current_user, load_user=main.load_user,
                          insert_account=main.insert_account, welcome_account=main.welcome_created_account)
from .medical_visits import register_medical_routes
register_medical_routes(app, get_db_connection=main.get_db_connection,
                        get_current_user=main.get_current_user, load_user=main.load_user, can_manage=main.user_can_manage)
register_resource_maintenance(app, get_db_connection=main.get_db_connection,
                             get_current_user=main.get_current_user, load_user=main.load_user)
# Production runs one API and owns its workers. Tests/candidates can disable mail.
if os.getenv('PRESENCE_CONTRACT_REMINDER_WORKER', '1') == '0':
    main.start_contract_reminder_worker = lambda: None
register_attendance_routes(
    app, get_db_connection=main.get_db_connection, get_current_user=main.get_current_user,
    load_user=main.load_user, can_manage=main.user_can_manage,
    save_config=main.save_config, save_checked=main.save_config_checked,
)


def readiness():
    return server_signing_readiness(os.getenv('CLAWSHOW_ESIGN_API_KEY', ''), os.getenv('PRESENCE_ESIGN_PUBLIC_URL', ''))


attempts_path = Path(os.getenv('PRESENCE_SIGNATURE_ATTEMPTS_DB', '/var/lib/presence-app/signature-attempts.sqlite3'))
register_profile_routes(app, get_db_connection=main.get_db_connection, get_current_user=main.get_current_user,
                        load_user=main.load_user, attempts=SignatureAttempts(attempts_path))
register_signature_preflight(
    app, get_current_user=main.get_current_user, get_db_connection=main.get_db_connection,
    load_user=main.load_user, require_target_access=main.require_target_access,
    get_config=main.get_config, readiness=readiness, sender=main.send_fiche_for_signature,
    attempts=SignatureAttempts(attempts_path),
)
