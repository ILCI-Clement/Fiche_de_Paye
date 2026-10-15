import unittest
from backend.server_signing_readiness import server_signing_readiness


class ServerSigningReadinessTests(unittest.TestCase):
    def test_missing_key_is_disabled(self):
        self.assertFalse(server_signing_readiness('', 'https://pdf.example.test')['can_send'])

    def test_sandbox_is_not_production(self):
        self.assertEqual(server_signing_readiness('sk_test_fixture', 'https://pdf.example.test')['mode'], 'sandbox-only')

    def test_live_warns_about_real_mail(self):
        result = server_signing_readiness('sk_live_fixture', 'https://pdf.example.test')
        self.assertTrue(result['can_send'])
        self.assertEqual(result['mode'], 'production')
        self.assertIn('e-mails', result['reason'])
        self.assertNotIn('sk_live_fixture', str(result))

    def test_invalid_public_addresses_fail_closed(self):
        for url in ['http://pdf.example.test', 'https://localhost', 'https://127.0.0.1', 'https://[::1]',
                    'https://user:password@pdf.example.test', 'https://pdf.example.test?secret=1',
                    'https://pdf.example.test#fragment', 'https://pdf.example.test:bad', 'https://[invalid']:
            with self.subTest(url=url):
                self.assertFalse(server_signing_readiness('sk_test_fixture', url)['can_send'])


if __name__ == '__main__':
    unittest.main()
