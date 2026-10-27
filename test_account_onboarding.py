"""Initial credentials and welcome delivery on synthetic accounts, SMTP mocked."""
import json
import smtplib
import unittest
from unittest.mock import Mock, patch

from backend.account_onboarding import credentials_for_account, send_welcome_email
import test_super_admin_roles as fixtures


class InitialCredentialTests(unittest.TestCase):
    def test_generated_passwords_have_exact_agreed_alphabet_and_no_plaintext_repr(self):
        for _ in range(100):
            result = credentials_for_account({'new_mail': 'synthetic@example.com', 'generate_password': True})
            value = result.password.get_secret_value()
            self.assertRegex(value, r'^[A-Z0-9]{8}$')
            self.assertRegex(value, '[A-Z]')
            self.assertRegex(value, '[0-9]')
            self.assertNotIn(value, repr(result))

    def test_invalid_email_modes_and_bcrypt_byte_limits_fail_before_creation(self):
        base = {'new_mail': 'synthetic@example.com', 'new_password': 'synthetic-password'}
        for change in ({'new_mail': ''}, {'new_mail': 'invalid'}, {'new_mail': 'a@example.com\r\nBcc:other@example.com'},
                       {'generate_password': 'true'}, {'generate_password': True},
                       {'new_password': 'short'}, {'new_password': 'a' * 73}, {'new_password': 'é' * 37}):
            with self.subTest(change=change), self.assertRaises(ValueError):
                credentials_for_account({**base, **change})
        self.assertEqual(credentials_for_account({**base, 'new_password': 'a' * 72}).password.get_secret_value(), 'a' * 72)


class WelcomeTransportTests(unittest.TestCase):
    def send(self, **changes):
        return send_welcome_email(email='synthetic@example.com', name='Synthetic <Name>',
            password='SYNTHETIC123', generated=True, app_url='https://example.com/',
            server='smtp.example.test', port=changes.pop('port', 587), sender='sender@example.com',
            smtp_password='synthetic-smtp', **changes)

    def test_starttls_certificate_validation_and_authentication_precede_send(self):
        with patch('backend.account_onboarding.smtplib.SMTP') as factory:
            transport = factory.return_value.__enter__.return_value
            transport.sendmail.return_value = {}
            self.assertTrue(self.send())
            calls = [entry[0] for entry in transport.method_calls]
            self.assertLess(calls.index('starttls'), calls.index('login'))
            self.assertLess(calls.index('login'), calls.index('sendmail'))
            self.assertTrue(transport.starttls.call_args.kwargs['context'].check_hostname)
            transport.sendmail.assert_called_once()

    def test_ssl_port_and_failed_delivery_return_status_without_retry(self):
        with patch('backend.account_onboarding.smtplib.SMTP_SSL') as factory:
            transport = factory.return_value.__enter__.return_value
            transport.sendmail.side_effect = smtplib.SMTPException('Synthetic failure')
            self.assertFalse(self.send(port=465))
            self.assertTrue(factory.call_args.kwargs['context'].check_hostname)
            transport.starttls.assert_not_called()
            transport.sendmail.assert_called_once()


class WelcomeAccountTransactionTests(unittest.TestCase):
    setUp = fixtures.MainRoleMutationTests.setUp

    def payload(self):
        return {'first_name': 'Synthetic', 'last_name': 'Onboarding', 'new_mail': 'new@example.com',
                'generate_password': True, 'role_tags': ['Employe'], 'manager_id': 'admin'}

    def test_generated_credentials_are_hashed_and_delivered_only_once_after_commit(self):
        observed = {}
        def hashed(value):
            observed['password'] = value
            return 'synthetic-hash-only'
        def delivery(email, name, password, generated):
            self.assertEqual(self.fixture.commits, 1)
            self.assertEqual(password, observed['password'])
            self.assertTrue(generated)
            return True
        with patch.object(fixtures.main, 'hash_password', side_effect=hashed), patch.object(
                fixtures.main, 'send_account_welcome_email', side_effect=delivery) as mail:
            result = fixtures.main.create_user(self.payload(), self.fixture.actor())
        created, = [row for key, row in self.fixture.users.items() if key not in {'root', 'admin', 'worker'}]
        self.assertTrue(created['initial_password_pending'])
        self.assertEqual(created['password_hash'], 'synthetic-hash-only')
        self.assertNotIn(observed['password'], json.dumps(result))
        self.assertTrue(result['welcome_email_sent'])
        mail.assert_called_once()

    def test_email_failure_keeps_committed_account_and_reports_no_delivery(self):
        with patch.object(fixtures.main, 'send_account_welcome_email', side_effect=OSError('Synthetic failure')) as mail:
            result = fixtures.main.create_user(self.payload(), self.fixture.actor())
        self.assertEqual(result['status'], 'success')
        self.assertFalse(result['welcome_email_sent'])
        self.assertEqual(len(self.fixture.users), 4)
        self.assertEqual(self.fixture.rollbacks, 0)
        mail.assert_called_once()

    def test_transaction_failure_and_duplicate_email_never_send(self):
        with patch.object(fixtures.main, 'send_account_welcome_email') as mail:
            self.fixture.fail_audit = True
            with self.assertRaises(RuntimeError):
                fixtures.main.create_user(self.payload(), self.fixture.actor())
            self.assertEqual(len(self.fixture.users), 3)
            mail.assert_not_called()
            self.fixture.fail_audit = False
            self.fixture.users['worker']['email'] = 'NEW@example.com'
            with self.assertRaises(fixtures.HTTPException) as caught:
                fixtures.main.create_user(self.payload(), self.fixture.actor())
            self.assertEqual(caught.exception.status_code, 400)
            mail.assert_not_called()


if __name__ == '__main__':
    unittest.main()
