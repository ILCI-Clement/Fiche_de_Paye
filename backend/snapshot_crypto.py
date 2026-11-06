"""Separate snapshot key ring and purpose-bound AES-256-GCM envelopes."""
import json
import os
import secrets
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from .personnel_sensitive import _load_keys, _encode

MAX_SNAPSHOT_BYTES = 24 * 1024 * 1024


def keys():
    path = os.getenv('PRESENCE_PERSONNEL_SNAPSHOT_KEYS_FILE')
    if not path:
        raise ValueError('snapshot_key_unavailable')
    return _load_keys(path)


def protect(content, identity, variant):
    if len(content) > MAX_SNAPSHOT_BYTES:
        raise ValueError('snapshot_too_large')
    ring = keys()
    nonce = secrets.token_bytes(12)
    aad = f'rh.register.snapshot.v1:{identity}:{variant}:{ring.active_key_id}'.encode()
    encrypted = AESGCM(ring.keys[ring.active_key_id]).encrypt(nonce, content, aad)
    header = json.dumps({'version': 1, 'key_id': ring.active_key_id, 'nonce': _encode(nonce)}, separators=(',', ':')).encode()
    return len(header).to_bytes(4, 'big') + header + encrypted


def reveal(envelope, identity, variant):
    if not isinstance(envelope, bytes) or not 20 < len(envelope) < MAX_SNAPSHOT_BYTES + 4096:
        raise ValueError('snapshot_corrupt')
    size = int.from_bytes(envelope[:4], 'big')
    if not 10 <= size <= 2048:
        raise ValueError('snapshot_corrupt')
    header = json.loads(envelope[4:4 + size])
    if set(header) != {'version', 'key_id', 'nonce'} or header['version'] != 1:
        raise ValueError('snapshot_corrupt')
    from .personnel_sensitive import _decode
    nonce = _decode(header['nonce'], 12)
    aad = f"rh.register.snapshot.v1:{identity}:{variant}:{header['key_id']}".encode()
    content = AESGCM(keys().keys[header['key_id']]).decrypt(nonce, envelope[4 + size:], aad)
    if len(content) > MAX_SNAPSHOT_BYTES:
        raise ValueError('snapshot_too_large')
    return content
