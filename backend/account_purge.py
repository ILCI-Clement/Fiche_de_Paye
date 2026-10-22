"""Super-admin account removal after its business dependencies have been resolved.

Do not cascade across another person's documents or attendance workspaces. Immutable
actor references in event ledgers remain as historical identifiers after removal.
The caller locks all users and validates a fresh super-admin and confirmation first.
"""
import json
from fastapi import HTTPException
if __package__:
    from .destructive_audit import append_deletion_event
    from .role_policy import is_super_admin, reserve_deleted_account_name
else:
    from destructive_audit import append_deletion_event
    from role_policy import is_super_admin, reserve_deleted_account_name


def purge_account(cursor, actor, target, *, reason):
    if not is_super_admin(actor):
        raise HTTPException(403, 'La suppression définitive est réservée au Super Admin.')
    username = target['username']
    if username == actor['username'] or target.get('is_active', True):
        raise HTTPException(409, 'Désactivez d’abord un autre compte avant sa suppression définitive.')
    if not isinstance(reason, str) or not reason.strip() or len(reason) > 500:
        raise HTTPException(400, 'Indiquez un motif de confirmation (500 caractères maximum).')
    dependencies = {}
    for table, column, label in (
        ('users', 'manager_username', 'personnes_rattachées'),
        ('personnel_contracts', 'username', 'contrats'),
        ('transport_receipts', 'employee_username', 'documents'),
        ('leave_requests', 'employee_username', 'congés'),
        ('annual_interviews', 'employee_username', 'entretiens'),
        ('personnel_register', 'linked_username', 'registre'),
    ):
        cursor.execute(f'SELECT COUNT(*) AS total FROM {table} WHERE {column}=%s', (username,))
        count = cursor.fetchone()['total']
        if count:
            dependencies[label] = int(count)
    cursor.execute('SELECT user_id,form_content FROM Presence FOR UPDATE')
    empty_owned = False
    for row in cursor.fetchall():
        try:
            config = row['form_content']
            if isinstance(config, str):
                config = json.loads(config)
            records = config.get('employes_data', [])
            if not isinstance(records, list) or any(not isinstance(record, dict) for record in records):
                raise ValueError()
        except (ValueError, TypeError, AttributeError):
            # An unreadable workspace cannot safely be declared unrelated.
            raise HTTPException(409, 'Une configuration de présence doit être vérifiée avant la suppression.') from None
        if row['user_id'] == username:
            if records:
                dependencies['fiches_dans_espace'] = len(records)
            else:
                empty_owned = True
        email = str(target.get('email') or '').strip().casefold()
        linked = sum(record.get('account_username') == username or (
            not record.get('account_username') and bool(email)
            and str(record.get('email_employe') or '').strip().casefold() == email
        ) for record in records)
        if linked:
            dependencies['fiches_liées'] = dependencies.get('fiches_liées', 0) + linked
    if dependencies:
        raise HTTPException(409, {
            'message': 'Des données sont encore rattachées à ce compte. Traitez-les séparément avant sa suppression définitive.',
            'dependencies': dependencies,
        })
    cursor.execute('SELECT COUNT(*) AS total FROM user_group_memberships WHERE username=%s', (username,))
    memberships = int(cursor.fetchone()['total'])
    cursor.execute('DELETE FROM user_group_memberships WHERE username=%s', (username,))
    if empty_owned:
        cursor.execute('DELETE FROM Presence WHERE user_id=%s', (username,))
    # Sent-reminder rows are operational dispatch history, not a cascade into files.
    # Keep them and all immutable event actors; they do not authorize a login.
    reserve_deleted_account_name(cursor, target)
    cursor.execute('DELETE FROM users WHERE username=%s', (username,))
    append_deletion_event(cursor, actor, 'account', username,
                          {'deleted_records': 1, 'membership_count': memberships, 'was_active': False})
    return {'deleted_records': 1, 'membership_count': memberships}
