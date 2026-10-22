"""Field encryption for personnel NIR values; configuration never creates a production key.

The key file must be kept outside source control and ordinary database backups. Retain
old keys when rotating the active key. Encryption authenticates the employee identity,
key id, format version and purpose; moving ciphertext between employees fails closed.
"""
import base64
from dataclasses import dataclass
import json
import os
from pathlib import Path
import re
import secrets
from types import MappingProxyType
from typing import Mapping

from cryptography.exceptions import InvalidTag
from cryptography.hazmat.primitives.ciphers.aead import AESGCM


_KEY_FILE_ENV = 'PRESENCE_PERSONNEL_NIR_KEYS_FILE'
_MAX_KEY_FILE_BYTES = 16_384
_MAX_ENVELOPE_BYTES = 2048
_MAX_KEYS = 32
_MAX_EMPLOYEE_ID = 2**63 - 1
_KEY_ID = re.compile(r'[A-Za-z0-9][A-Za-z0-9_-]{0,63}\Z', re.ASCII)
_BASE64URL = re.compile(r'[A-Za-z0-9_-]+={0,2}\Z', re.ASCII)
# Administrative birth-month codes 20..99 are accepted. The check key is 01..97.
# This checks syntax only, neither the modulo-97 identity key nor a real person's identity.
_NIR = re.compile(
    r'[1278][0-9]{2}(?:0[1-9]|1[0-2]|[2-9][0-9])'
    r'(?:[0-9]{2}|2[AB])[0-9]{6}(?:0[1-9]|[1-8][0-9]|9[0-7])\Z', re.ASCII,
)
_VERSION = 1
_ALGORITHM = 'A256GCM'
_PURPOSE = 'presence.personnel.social_security_number'
_ENVELOPE_FIELDS = frozenset({'version', 'algorithm', 'key_id', 'nonce', 'ciphertext'})
_PROTECTION_MESSAGE = 'La protection du numéro de sécurité sociale est indisponible.'


class NirProtectionError(Exception):
    """An intentionally generic error: never carry a value, key, envelope or path."""
    def __init__(self):
        super().__init__(_PROTECTION_MESSAGE)


@dataclass(frozen=True, repr=False)
class _KeyRing:
    active_key_id: str
    keys: Mapping[str, bytes]

    def __repr__(self):
        return '<NirKeyRing protected>'


def normalize_social_security_number(value: str | None) -> str | None:
    """Normalize user formatting, without pretending to validate an actual identity."""
    if value is None:
        return None
    if not isinstance(value, str) or len(value) > 128:
        raise NirProtectionError()
    if not value.strip():
        return None
    normalized = re.sub(r'[\s-]+', '', value).upper()
    if not _NIR.fullmatch(normalized):
        raise NirProtectionError()
    return normalized


def _pairs_without_duplicates(pairs):
    result = {}
    for key, value in pairs:
        if key in result:
            raise NirProtectionError()
        result[key] = value
    return result


def _strict_json(value):
    try:
        return json.loads(value, object_pairs_hook=_pairs_without_duplicates,
                          parse_constant=lambda _: _reject_json_constant())
    except (ValueError, TypeError, RecursionError):
        raise NirProtectionError() from None


def _reject_json_constant():
    raise NirProtectionError()


def _encode(value):
    return base64.urlsafe_b64encode(value).decode('ascii').rstrip('=')


def _decode(value, expected_bytes):
    if not isinstance(value, str) or not _BASE64URL.fullmatch(value):
        raise NirProtectionError()
    # Bound allocation before decoding and require canonical base64url spelling.
    canonical_length = (expected_bytes * 8 + 5) // 6
    if len(value.rstrip('=')) != canonical_length or len(value) > canonical_length + 2:
        raise NirProtectionError()
    try:
        decoded = base64.b64decode(value + '=' * (-len(value) % 4), altchars=b'-_', validate=True)
    except (ValueError, TypeError):
        raise NirProtectionError() from None
    if len(decoded) != expected_bytes or _encode(decoded) != value.rstrip('='):
        raise NirProtectionError()
    return decoded


def _load_keys(path=None):
    configured = path if path is not None else os.getenv(_KEY_FILE_ENV)
    if not configured:
        raise NirProtectionError()
    try:
        with Path(configured).open('rb') as stream:
            content = stream.read(_MAX_KEY_FILE_BYTES + 1)
        if len(content) > _MAX_KEY_FILE_BYTES:
            raise NirProtectionError()
        document = _strict_json(content.decode('utf-8-sig'))
    except (OSError, ValueError, TypeError, UnicodeError):
        raise NirProtectionError() from None
    if not isinstance(document, dict) or set(document) != {'active_key_id', 'keys'}:
        raise NirProtectionError()
    active, encoded = document['active_key_id'], document['keys']
    if (not isinstance(active, str) or not _KEY_ID.fullmatch(active)
            or not isinstance(encoded, dict) or not 1 <= len(encoded) <= _MAX_KEYS
            or active not in encoded):
        raise NirProtectionError()
    keys = {}
    for key_id, value in encoded.items():
        if not isinstance(key_id, str) or not _KEY_ID.fullmatch(key_id):
            raise NirProtectionError()
        keys[key_id] = _decode(value, 32)
    return _KeyRing(active, MappingProxyType(keys))


def nir_storage_ready() -> bool:
    """Only report whether independent key material is valid; never create or log it."""
    try:
        _load_keys()
        return True
    except NirProtectionError:
        return False


def _aad(employee_id, key_id):
    if type(employee_id) is not int or not 1 <= employee_id <= _MAX_EMPLOYEE_ID:
        raise NirProtectionError()
    return json.dumps({'purpose': _PURPOSE, 'version': _VERSION,
                       'key_id': key_id, 'employee_id': employee_id},
                      sort_keys=True, separators=(',', ':')).encode('ascii')


def encrypt_social_security_number(value: str, employee_id: int) -> str:
    try:
        normalized = normalize_social_security_number(value)
        if normalized is None:
            raise NirProtectionError()
        ring = _load_keys()
        key_id = ring.active_key_id
        nonce = secrets.token_bytes(12)
        ciphertext = AESGCM(ring.keys[key_id]).encrypt(nonce, normalized.encode('ascii'), _aad(employee_id, key_id))
        return json.dumps({'version': _VERSION, 'algorithm': _ALGORITHM, 'key_id': key_id,
                           'nonce': _encode(nonce), 'ciphertext': _encode(ciphertext)},
                          sort_keys=True, separators=(',', ':'))
    except (ValueError, TypeError, UnicodeError, OverflowError):
        raise NirProtectionError() from None


def decrypt_social_security_number(envelope: str, employee_id: int) -> str:
    try:
        if (not isinstance(envelope, str) or len(envelope) > _MAX_ENVELOPE_BYTES
                or len(envelope.encode('utf-8')) > _MAX_ENVELOPE_BYTES):
            raise NirProtectionError()
        document = _strict_json(envelope)
        if (not isinstance(document, dict) or set(document) != _ENVELOPE_FIELDS
                or type(document['version']) is not int or document['version'] != _VERSION
                or document['algorithm'] != _ALGORITHM
                or not isinstance(document['key_id'], str) or not _KEY_ID.fullmatch(document['key_id'])):
            raise NirProtectionError()
        key_id = document['key_id']
        ring = _load_keys()
        if key_id not in ring.keys:
            raise NirProtectionError()
        nonce = _decode(document['nonce'], 12)
        ciphertext = _decode(document['ciphertext'], 31)  # 15 ASCII characters + GCM's 16-byte tag.
        plaintext = AESGCM(ring.keys[key_id]).decrypt(nonce, ciphertext, _aad(employee_id, key_id)).decode('ascii')
        if normalize_social_security_number(plaintext) != plaintext:
            raise NirProtectionError()
        return plaintext
    except (InvalidTag, ValueError, TypeError, UnicodeError, OverflowError):
        raise NirProtectionError() from None


def ensure_local_nir_key_file(path):
    """Explicit local setup only: generate once with exclusive creation, never overwrite.

    POSIX mode is 0600. On Windows the containing private directory must also have an
    appropriate NTFS ACL. Production loading never calls this provisioning helper.
    """
    try:
        target = Path(path)
        if target.exists():
            _load_keys(target)
            return
    except (OSError, ValueError, TypeError):
        raise NirProtectionError() from None
    key_id = 'local-' + secrets.token_hex(8)
    encoded = json.dumps({'active_key_id': key_id, 'keys': {key_id: _encode(secrets.token_bytes(32))}},
                         sort_keys=True, separators=(',', ':')).encode('utf-8')
    descriptor = None
    created = False
    try:
        # Parent creation is the caller's job so this helper cannot choose another location.
        descriptor = os.open(target, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
        created = True
        with os.fdopen(descriptor, 'wb') as stream:
            descriptor = None
            stream.write(encoded)
            stream.flush()
            os.fsync(stream.fileno())
    except FileExistsError:
        _load_keys(target)
    except (OSError, ValueError, TypeError):
        if descriptor is not None:
            os.close(descriptor)
        if created:
            try:
                target.unlink()
            except OSError:
                pass
        raise NirProtectionError() from None
