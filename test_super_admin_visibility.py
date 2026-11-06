"""Synthetic identities only; system identity hiding never removes business history."""
import json
import unittest
from unittest.mock import MagicMock, patch

from backend.person_visibility import account_list_visible
from test_super_admin_roles import account
from backend import main
import test_team_directory as directory_fixture
from test_team_directory import GROUPS
import test_account_personnel as register_fixture


class VisibilityPolicyTests(unittest.TestCase):
    def test_only_owner_can_see_super_admin_even_other_super_admin_is_denied(self):
        root = account('root', ['SuperAdmin'])
        for role in ['Employe', 'Responsable', 'Admin', 'SuperAdmin']:
            self.assertFalse(account_list_visible(root, account('other', [role])))
        self.assertTrue(account_list_visible(root, root))
        self.assertTrue(account_list_visible(account('manager', ['Responsable']), root))
        self.assertFalse(account_list_visible(None, root))
        self.assertFalse(account_list_visible({'username': 'root', 'role_tags': json.dumps(['SuperAdmin'])}, account('a', ['Admin'])))

    def test_account_catalog_filters_before_public_projection_including_inactive_tab(self):
        users = [account('root', ['SuperAdmin']), account('root2', ['SuperAdmin']), account('admin', ['Admin']), account('worker', ['Employe'])]
        connection = MagicMock()
        with patch.object(main, 'get_db_connection', return_value=connection), patch.object(main, 'all_users', return_value=users), patch.object(main, 'admin_job_titles', return_value={}):
            for actor, expected in [(users[0], ['root', 'admin', 'worker']), (users[2], ['admin', 'worker'])]:
                for inactive in [True, False]:
                    self.assertEqual([u['username'] for u in main.list_users(actor, inactive)['users']], expected)


class DirectoryVisibilityTests(unittest.TestCase):
    setUp = directory_fixture.DirectoryApiTests.setUp

    def test_directory_and_team_hide_other_system_accounts_but_keep_own(self):
        for role in ['Admin', 'Employe', 'SuperAdmin']:
            self.actor = {'username': 'viewer', 'role_tags': [role]}
            rows = [{'username': 'hidden', 'email': 'hidden@example.test', 'role_tags': '["SuperAdmin"]'},
                    {'username': 'viewer', 'email': 'viewer@example.test', 'role_tags': json.dumps([role])}]
            self.cursor.fetchall.side_effect = [GROUPS, rows, [], []]
            result = self.client.get('/organization-directory')
            self.assertEqual(result.status_code, 200)
            self.assertEqual([u['username'] for u in result.json()['users']], ['viewer'])
            links = [dict(row, group_id=2, relation='member') for row in rows]
            self.cursor.fetchall.side_effect = [[GROUPS[1]], links, []]
            result = self.client.get('/my-team')
            self.assertEqual(len(result.json()['members']), 1)


class RegisterVisibilityTests(unittest.TestCase):
    setUp = register_fixture.AccountPersonnelHTTPTests.setUp
    create = register_fixture.AccountPersonnelHTTPTests.create
    rows = register_fixture.AccountPersonnelHTTPTests.rows
    raw_update = register_fixture.AccountPersonnelHTTPTests.raw_update

    def test_linked_system_identity_hidden_before_pagination_and_account_search(self):
        self.raw_update('UPDATE users SET role_tags=? WHERE username=?', ('["SuperAdmin","Admin"]', 'fixture-employee'))
        result = self.create(linked_username='fixture-employee')
        self.assertEqual(result.status_code, 201, result.text)
        for path in ['/personnel-register', '/personnel-register?q=fixture', '/personnel-register?archive=all']:
            result = self.client.get(path)
            self.assertEqual(result.status_code, 200, result.text)
            self.assertEqual(result.json()['total'], 0)
        result = self.client.get('/personnel-register/accounts')
        self.assertEqual(result.status_code, 200, result.text)
        self.assertNotIn('fixture-employee', [u['username'] for u in result.json()['accounts']])
        self.assertEqual(self.client.get('/account-personnel/fixture-employee').status_code, 403)
        self.actor = {'username': 'fixture-employee', 'role_tags': ['SuperAdmin', 'Admin'], 'session_version': 0}
        result = self.client.get('/personnel-register')
        self.assertEqual(result.status_code, 200, result.text)
        self.assertEqual(result.json()['total'], 1)
        self.assertEqual(len(self.rows('personnel_register')), 1)


if __name__ == '__main__':
    unittest.main()
