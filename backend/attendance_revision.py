"""Content revisions for optimistic attendance saves without changing stored JSON."""
import hashlib
import json
from fastapi import HTTPException


def config_revision(config):
    return hashlib.sha256(json.dumps(config, sort_keys=True, separators=(',', ':'),
                                     ensure_ascii=False, default=str).encode()).hexdigest()


def require_revision(config, expected):
    if expected is not None and config_revision(config) != expected:
        raise HTTPException(409, 'Cette configuration a été modifiée dans une autre fenêtre. '
                            'Votre brouillon est conservé. Rechargez la version du serveur avant de continuer.')
