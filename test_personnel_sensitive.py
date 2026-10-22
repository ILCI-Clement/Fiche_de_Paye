"""Synthetic NIR and temporary key-file tests; no real employee, database or service."""
import base64
import json
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from pydantic import SecretStr, ValidationError

from backend import personnel_sensitive as protected
from backend.personnel_rules import EmployeeCreate, EmployeeUpdate


# Intentionally synthetic: the order component 000 does not identify a real person.
SYNTHETIC_NIR = '199129900000001'
SYNTHETIC_CORSICA = '299122A00000001'
ENVIRONMENT_KEY = 'PRESENCE_PERSONNEL_NIR_KEYS_FILE'


def encode(value):
    return base64.urlsafe_b64encode(value).decode('ascii').rstrip('=')


class NirEncryptionTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory(prefix='presence-nir-unit-')
        self.addCleanup(self.directory.cleanup)
        self.key_path = Path(self.directory.name) / 'keys.json'
        protected.ensure_local_nir_key_file(self.key_path)
        self.environment = patch.dict(os.environ, {ENVIRONMENT_KEY: str(self.key_path)})
        self.environment.start()
        self.addCleanup(self.environment.stop)

    def encrypt(self, employee_id=21):
        return protected.encrypt_social_security_number(SYNTHETIC_NIR, employee_id)

    def write_keys(self, document):
        self.key_path.write_text(json.dumps(document), encoding='utf-8')

    def assert_protection_failure(self, operation):
        with self.assertRaises(protected.NirProtectionError) as caught:
            operation()
        self.assertEqual(str(caught.exception), protected._PROTECTION_MESSAGE)
        self.assertNotIn(SYNTHETIC_NIR, repr(caught.exception))
        self.assertNotIn(str(self.key_path), str(caught.exception))

    def test_aes_roundtrip_without_plaintext_in_envelope(self):
        envelope = self.encrypt()
        self.assertNotIn(SYNTHETIC_NIR, envelope)
        self.assertEqual(protected.decrypt_social_security_number(envelope, 21), SYNTHETIC_NIR)
        document = json.loads(envelope)
        self.assertEqual(set(document), protected._ENVELOPE_FIELDS)
        self.assertEqual(document['version'], 1)
        self.assertEqual(document['algorithm'], 'A256GCM')

    def test_random_nonce_makes_same_value_distinct(self):
        first, second = self.encrypt(), self.encrypt()
        self.assertNotEqual(first, second)
        self.assertNotEqual(json.loads(first)['nonce'], json.loads(second)['nonce'])

    def test_format_is_normalized_before_encryption(self):
        value = '2 99-12 2a 000 000 01'
        envelope = protected.encrypt_social_security_number(value, 21)
        self.assertEqual(protected.decrypt_social_security_number(envelope, 21), SYNTHETIC_CORSICA)

    def test_ciphertext_cannot_move_to_another_employee(self):
        envelope = self.encrypt()
        self.assert_protection_failure(lambda: protected.decrypt_social_security_number(envelope, 22))

    def test_purpose_is_authenticated(self):
        envelope = self.encrypt()
        with patch.object(protected, '_PURPOSE', 'a-different-purpose'):
            self.assert_protection_failure(lambda: protected.decrypt_social_security_number(envelope, 21))

    def test_employee_identifier_is_strict(self):
        for employee_id in (0, -1, True, '21', None, 2**63):
            with self.subTest(kind=type(employee_id).__name__):
                self.assert_protection_failure(lambda: self.encrypt(employee_id))

    def test_modified_nonce_or_ciphertext_fails_closed(self):
        original = json.loads(self.encrypt())
        for field, length in (('nonce', 12), ('ciphertext', 31)):
            changed = dict(original)
            raw = bytearray(protected._decode(changed[field], length))
            raw[0] ^= 1
            changed[field] = encode(bytes(raw))
            with self.subTest(field=field):
                self.assert_protection_failure(lambda: protected.decrypt_social_security_number(json.dumps(changed), 21))

    def test_version_algorithm_and_json_fields_are_strict(self):
        original = json.loads(self.encrypt())
        mutations = [
            {**original, 'version': 2}, {**original, 'version': True},
            {**original, 'algorithm': 'unsupported'}, {**original, 'extra': 'not-accepted'},
            {key: value for key, value in original.items() if key != 'nonce'},
        ]
        for changed in mutations:
            self.assert_protection_failure(lambda: protected.decrypt_social_security_number(json.dumps(changed), 21))

    def test_invalid_json_and_size_are_rejected(self):
        for envelope in ('', '[]', 'null', 'not json', '{"version":1,"version":1}',
                         '{"version":NaN}', ' ' * 2049, 123, None):
            self.assert_protection_failure(lambda: protected.decrypt_social_security_number(envelope, 21))

    def test_base64url_spelling_and_lengths_are_strict(self):
        original = json.loads(self.encrypt())
        for bad in ('', 'a', '*' * 16, 'a' * 2000, 123, None):
            changed = {**original, 'nonce': bad}
            self.assert_protection_failure(lambda: protected.decrypt_social_security_number(json.dumps(changed), 21))

    def test_rotated_active_key_can_read_retained_old_key(self):
        old_envelope = self.encrypt()
        document = json.loads(self.key_path.read_text(encoding='utf-8'))
        document['keys']['rotated'] = encode(os.urandom(32))
        document['active_key_id'] = 'rotated'
        self.write_keys(document)
        self.assertEqual(protected.decrypt_social_security_number(old_envelope, 21), SYNTHETIC_NIR)
        new_envelope = self.encrypt()
        self.assertEqual(json.loads(new_envelope)['key_id'], 'rotated')
        self.assertEqual(protected.decrypt_social_security_number(new_envelope, 21), SYNTHETIC_NIR)

    def test_key_id_is_authenticated_even_with_identical_key_bytes(self):
        document = json.loads(self.key_path.read_text(encoding='utf-8'))
        document['keys']['same-key-other-id'] = document['keys'][document['active_key_id']]
        self.write_keys(document)
        changed = json.loads(self.encrypt())
        changed['key_id'] = 'same-key-other-id'
        self.assert_protection_failure(lambda: protected.decrypt_social_security_number(json.dumps(changed), 21))

    def test_removed_or_unknown_key_fails_closed(self):
        envelope = self.encrypt()
        self.write_keys({'active_key_id': 'replacement', 'keys': {'replacement': encode(os.urandom(32))}})
        self.assert_protection_failure(lambda: protected.decrypt_social_security_number(envelope, 21))

    def test_missing_environment_does_not_generate_a_key(self):
        with patch.dict(os.environ, {}, clear=True):
            self.assertFalse(protected.nir_storage_ready())
            self.assert_protection_failure(lambda: self.encrypt())

    def test_missing_file_does_not_get_implicitly_created(self):
        missing = Path(self.directory.name) / 'not-created.json'
        with patch.dict(os.environ, {ENVIRONMENT_KEY: str(missing)}):
            self.assertFalse(protected.nir_storage_ready())
            self.assert_protection_failure(lambda: self.encrypt())
            self.assertFalse(missing.exists())

    def test_malformed_key_documents_fail_closed(self):
        valid = json.loads(self.key_path.read_text(encoding='utf-8'))
        cases = [[], {}, {**valid, 'extra': 'rejected'}, {'active_key_id': 'missing', 'keys': valid['keys']},
                 {'active_key_id': 'bad id', 'keys': {'bad id': encode(os.urandom(32))}},
                 {'active_key_id': 'x', 'keys': {'x': encode(os.urandom(16))}},
                 {'active_key_id': 'x', 'keys': {'x': 'a' * 10000}},
                 {'active_key_id': 'x', 'keys': {'x': 123}}]
        for document in cases:
            self.write_keys(document)
            self.assertFalse(protected.nir_storage_ready())
            self.assert_protection_failure(lambda: self.encrypt())

    def test_key_file_duplicate_json_fields_or_oversize_are_rejected(self):
        for content in ('{"active_key_id":"x","active_key_id":"x","keys":{}}',
                        '{"active_key_id":"x","keys":{"x":"a","x":"b"}}',
                        ' ' * (protected._MAX_KEY_FILE_BYTES + 1), '\ud800'):
            self.key_path.write_bytes(content.encode('utf-8', errors='surrogatepass'))
            self.assertFalse(protected.nir_storage_ready())

    def test_key_ring_repr_is_redacted(self):
        ring = protected._load_keys()
        self.assertEqual(repr(ring), '<NirKeyRing protected>')
        for material in ring.keys.values():
            self.assertNotIn(encode(material), repr(ring))
            self.assertNotIn(repr(material), repr(ring))

    def test_existing_key_file_is_never_replaced(self):
        first = self.key_path.read_bytes()
        protected.ensure_local_nir_key_file(self.key_path)
        self.assertEqual(first, self.key_path.read_bytes())

    def test_invalid_existing_key_file_is_preserved_and_rejected(self):
        self.key_path.write_text('invalid-existing-key-file', encoding='utf-8')
        first = self.key_path.read_bytes()
        self.assert_protection_failure(lambda: protected.ensure_local_nir_key_file(self.key_path))
        self.assertEqual(first, self.key_path.read_bytes())

    def test_failed_local_creation_removes_only_its_new_partial_file(self):
        target = Path(self.directory.name) / 'new-partial.json'
        with patch.object(protected.os, 'fsync', side_effect=OSError('synthetic-write-failure')):
            self.assert_protection_failure(lambda: protected.ensure_local_nir_key_file(target))
        self.assertFalse(target.exists())

    def test_missing_parent_is_not_created_by_provisioning(self):
        target = Path(self.directory.name) / 'missing-parent' / 'key.json'
        self.assert_protection_failure(lambda: protected.ensure_local_nir_key_file(target))
        self.assertFalse(target.parent.exists())


class NirInputTests(unittest.TestCase):
    def test_only_name_remains_required(self):
        model = EmployeeCreate(full_name='Synthetic example')
        self.assertIsNone(model.social_security_number)
        self.assertNotIn('social_security_number', model.model_fields_set)

    def test_missing_means_preserve_and_explicit_empty_means_clear(self):
        missing = EmployeeUpdate(full_name='Synthetic example', expected_version=1)
        self.assertNotIn('social_security_number', missing.model_fields_set)
        for clear in (None, '', ' \t\r\n', '\u00a0'):
            value = EmployeeUpdate(full_name='Synthetic example', expected_version=1, social_security_number=clear)
            self.assertIn('social_security_number', value.model_fields_set)
            self.assertIsNone(value.social_security_number)

    def test_accepted_nir_is_normalized_secret_and_excluded_from_dumps(self):
        value = EmployeeCreate(full_name='Synthetic example', social_security_number='2 99-12 2a 000 000 01')
        self.assertIsInstance(value.social_security_number, SecretStr)
        self.assertEqual(value.social_security_number.get_secret_value(), SYNTHETIC_CORSICA)
        self.assertNotIn('social_security_number', value.model_dump())
        self.assertNotIn('social_security_number', value.model_dump_json())
        self.assertEqual(value.model_dump(include={'social_security_number'}), {})
        self.assertNotIn(SYNTHETIC_CORSICA, repr(value))
        self.assertNotIn('social_security_number', repr(value))

    def test_corsica_b_and_temporary_number_are_syntactically_supported(self):
        value = '799122B00000001'
        self.assertEqual(protected.normalize_social_security_number(value), value)

    def test_invalid_inputs_do_not_echo_numbers_in_validation_errors(self):
        for invalid in (SYNTHETIC_NIR + '0', '9' + SYNTHETIC_NIR[1:], SYNTHETIC_NIR[:-2] + '00',
                        '-' * 15, '1' * 129, int(SYNTHETIC_NIR), {'invalid': SYNTHETIC_NIR}):
            with self.assertRaises(ValidationError) as caught:
                EmployeeCreate(full_name='Synthetic example', social_security_number=invalid)
            text = str(caught.exception) + repr(caught.exception.errors()) + caught.exception.json()
            self.assertNotIn(SYNTHETIC_NIR, text)
            self.assertNotIn(str(invalid), text)

    def test_other_field_error_does_not_expose_valid_nir(self):
        with self.assertRaises(ValidationError) as caught:
            EmployeeCreate(full_name='', social_security_number=SYNTHETIC_NIR)
        self.assertNotIn(SYNTHETIC_NIR, str(caught.exception) + repr(caught.exception.errors()) + caught.exception.json())

    def test_format_validation_does_not_claim_checksum_authenticity(self):
        for ending in ('01', '02'):
            value = SYNTHETIC_NIR[:-2] + ending
            self.assertEqual(protected.normalize_social_security_number(value), value)


if __name__ == '__main__':
    unittest.main()
