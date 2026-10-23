"""Deterministic cooldown tests on disposable account fixtures, never real users."""
from datetime import datetime, timedelta
import unittest
from unittest.mock import patch

import test_account_lifecycle as fixture


class LoginGuardTests(unittest.TestCase):
    connect = fixture.AccountLifecycleTests.connect
    load_user = fixture.AccountLifecycleTests.load_user
    stored = fixture.AccountLifecycleTests.stored
    login = fixture.AccountLifecycleTests.login

    def setUp(self):
        fixture.AccountLifecycleTests.setUp(self)
        self.now = datetime(2026, 10, 23, 12)
        original = self.connect
        def connect():
            connection = original()
            connection.db.create_function('UTC_TIMESTAMP', 1, lambda _: self.now.isoformat(' '))
            return connection
        patch('backend.main.get_db_connection', connect).start()

    def wrong(self, identifier='employee'):
        return self.client.post('/login', json={'username': identifier, 'password': 'wrong synthetic password'})

    def test_fifth_failure_locks_aliases_for_ten_minutes(self):
        for _ in range(4):
            self.assertEqual(self.wrong().status_code, 401)
        fifth = self.wrong(' EMPLOYEE@EXAMPLE.TEST ')
        self.assertEqual(fifth.status_code, 429)
        self.assertEqual(fifth.headers['retry-after'], '600')
        self.now += timedelta(seconds=150)
        locked = self.login('employee@example.test')
        self.assertEqual(locked.status_code, 429)
        self.assertEqual(locked.headers['retry-after'], '450')
        self.assertEqual(self.stored()['login_failed_attempts'], 5)
        self.now += timedelta(seconds=450)
        self.assertEqual(self.login().status_code, 200)
        self.assertEqual(self.stored()['login_failed_attempts'], 0)

    def test_inactivity_reset_is_measured_from_last_failure(self):
        for _ in range(3): self.assertEqual(self.wrong().status_code, 401)
        self.now += timedelta(seconds=599)
        self.assertEqual(self.wrong().status_code, 401)
        self.assertEqual(self.stored()['login_failed_attempts'], 4)
        self.now += timedelta(seconds=600)
        self.assertEqual(self.wrong().status_code, 401)
        self.assertEqual(self.stored()['login_failed_attempts'], 1)

    def test_success_resets_failure_sequence(self):
        for _ in range(3): self.wrong()
        self.assertEqual(self.login().status_code, 200)
        self.assertEqual(self.wrong().status_code, 401)
        self.assertEqual(self.stored()['login_failed_attempts'], 1)

    def test_expired_cooldown_next_failure_starts_at_one(self):
        for _ in range(5): self.wrong()
        self.now += timedelta(seconds=600)
        self.assertEqual(self.wrong().status_code, 401)
        self.assertEqual(self.stored()['login_failed_attempts'], 1)

    def test_other_account_and_unknown_identifier_do_not_share_counter(self):
        for _ in range(5): self.wrong()
        self.assertEqual(self.login('manager').status_code, 200)
        for _ in range(6): self.assertEqual(self.wrong('missing@example.test').status_code, 401)
        self.assertEqual(self.stored('manager')['login_failed_attempts'], 0)
