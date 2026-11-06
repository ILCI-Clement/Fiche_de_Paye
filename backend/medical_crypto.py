"""Private PDF envelopes; keys live outside the repository and DB backups."""
import json
import os
import secrets
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from .personnel_sensitive import _load_keys, _encode, _decode

MAX_PDF_BYTES = 10 * 1024 * 1024


def keyring():
    path = os.getenv('PRESENCE_MEDICAL_KEYS_FILE')
    if not path:
        raise ValueError('medical_keys_unavailable')
    return _load_keys(path)


def protect(content, identity):
    ring = keyring()
    nonce = secrets.token_bytes(12)
    header = json.dumps({'key_id': ring.active_key_id, 'nonce': _encode(nonce)}, separators=(',', ':')).encode()
    aad = f'rh.medical.pdf.v1:{identity}:{ring.active_key_id}'.encode()
    return len(header).to_bytes(4, 'big') + header + AESGCM(ring.keys[ring.active_key_id]).encrypt(nonce, content, aad)


def reveal(envelope, identity):
    if not 20 < len(envelope) < MAX_PDF_BYTES + 4096:
        raise ValueError('medical_file_invalid')
    size = int.from_bytes(envelope[:4], 'big')
    if not 10 <= size <= 2048:
        raise ValueError('medical_file_invalid')
    header = json.loads(envelope[4:4 + size])
    aad = f"rh.medical.pdf.v1:{identity}:{header['key_id']}".encode()
    return AESGCM(keyring().keys[header['key_id']]).decrypt(_decode(header['nonce'], 12), envelope[4 + size:], aad)
