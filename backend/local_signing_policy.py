"""Explicit, expiring live-mail test permission for one approved signer pair."""
import time
from fastapi import HTTPException


class LocalSigningPolicy:
    def __init__(self, config, clock=time.time):
        self.config = config
        self.clock = clock

    @property
    def live(self):
        return bool(self.config.get('live_signing_test', {}).get('enabled'))

    @property
    def mode(self):
        return 'production-test' if self.live else 'sandbox-only'

    def check(self, employee_email=None, manager_email=None):
        if not self.live:
            return
        permission = self.config['live_signing_test']
        try:
            expires = float(permission.get('expires_at', 0))
        except (TypeError, ValueError):
            expires = 0
        if not 0 < expires - self.clock() <= 2 * 60 * 60 + 60:
            raise HTTPException(403, 'Le test d’e-mails réels est expiré ou invalide. Repassez en sandbox.')
        approved_employee = str(permission.get('employee_email') or '').strip().casefold()
        approved_manager = str(permission.get('manager_email') or '').strip().casefold()
        if not approved_employee or not approved_manager or approved_employee == approved_manager:
            raise HTTPException(403, 'La paire de signataires de test n’est pas configurée.')
        if employee_email is not None and (str(employee_email).strip().casefold() != approved_employee or
                                         str(manager_email or '').strip().casefold() != approved_manager):
            raise HTTPException(403, 'Ce test autorise uniquement la paire de signataires validée. Aucun autre destinataire ne peut être contacté.')

    def key(self):
        if self.live:
            try:
                self.check()
            except HTTPException:
                return ''  # Expired permission disables signing, not the rest of the app.
            key = str(self.config.get('live_test_key') or '')
            if not key.startswith('sk_live_'):
                raise RuntimeError('An explicitly configured live test key is required.')
            return key
        key = str(self.config.get('sandbox_key') or '')
        if key and not key.startswith('sk_test_'):
            raise RuntimeError('Invalid sandbox key prefix.')
        return key
