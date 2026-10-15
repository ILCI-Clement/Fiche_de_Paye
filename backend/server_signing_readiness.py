"""Pure configuration checks for the staged server entry; never sends a request."""
from urllib.parse import urlsplit


def server_signing_readiness(key: str, public_url: str):
    mode = 'sandbox-only' if key.startswith('sk_test_') else 'production'
    if not key.startswith(('sk_test_', 'sk_live_')):
        return {'mode': mode, 'can_send': False, 'reason': 'La clé de signature du serveur n’est pas configurée.'}
    try:
        parsed = urlsplit(public_url)
        valid = (parsed.scheme == 'https' and bool(parsed.hostname)
                 and parsed.hostname not in {'localhost', '127.0.0.1', '::1'}
                 and not parsed.username and not parsed.password and not parsed.query and not parsed.fragment)
        parsed.port  # Reject malformed ports without exposing the configured URL.
    except ValueError:
        valid = False
    if not valid:
        return {'mode': mode, 'can_send': False, 'reason': 'Configurez l’adresse HTTPS publique du service PDF.'}
    reason = ('Sandbox uniquement : aucun e-mail envoyé, documents SPECIMEN sans valeur juridique.'
              if mode == 'sandbox-only' else
              'Production : de vrais e-mails seront envoyés au salarié puis à son responsable direct. La demande consommera le quota de signature.')
    return {'mode': mode, 'can_send': True, 'reason': reason}
