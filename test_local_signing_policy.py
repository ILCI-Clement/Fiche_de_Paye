import unittest
from fastapi import HTTPException
from backend.local_signing_policy import LocalSigningPolicy


class LocalSigningPolicyTests(unittest.TestCase):
    def policy(self, **overrides):
        config = {'sandbox_key': 'sk_test_unit', 'live_test_key': 'sk_live_unit',
                  'live_signing_test': {'enabled': True, 'employee_email': 'e@example.com',
                                        'manager_email': 'm@example.com', 'expires_at': 2000}}
        config['live_signing_test'].update(overrides)
        return LocalSigningPolicy(config, clock=lambda: 1000)

    def test_sandbox_is_default(self):
        p = LocalSigningPolicy({'sandbox_key': 'sk_test_unit'})
        self.assertEqual(p.mode, 'sandbox-only'); self.assertEqual(p.key(), 'sk_test_unit')

    def test_live_only_accepts_approved_ordered_pair(self):
        p = self.policy()
        p.check('E@example.com', 'm@example.com')
        self.assertEqual(p.key(), 'sk_live_unit')
        for employee, manager in [('m@example.com','e@example.com'), ('other@example.com','m@example.com'), ('e@example.com','other@example.com')]:
            with self.assertRaises(HTTPException): p.check(employee, manager)

    def test_expiry_disables_live_signing_but_keeps_startup_available(self):
        p = self.policy(expires_at=999)
        with self.assertRaises(HTTPException): p.check('e@example.com','m@example.com')
        self.assertEqual(p.key(), '')

    def test_no_long_lived_permission(self):
        for expiry in [float('inf'), float('nan'), 1000000000]:
            with self.assertRaises(HTTPException): self.policy(expires_at=expiry).check()

    def test_missing_or_same_signer_fails(self):
        for override in [{'employee_email':''}, {'manager_email':'e@example.com'}]:
            with self.assertRaises(HTTPException): self.policy(**override).check()

    def test_production_key_is_not_implicitly_taken_from_sandbox_field(self):
        with self.assertRaises(RuntimeError): LocalSigningPolicy({'sandbox_key':'sk_live_unit'}).key()


if __name__ == '__main__': unittest.main()
