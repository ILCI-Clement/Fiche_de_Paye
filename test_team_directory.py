"""Read-only directory authorization and model regression, with no production writes."""
import unittest
from unittest.mock import MagicMock, patch
from fastapi import FastAPI, HTTPException
from fastapi.testclient import TestClient
from backend.team_directory import register_team_directory, directory_people, team_members
from organization_model import build_organization, person_label

GROUPS = [{'id': 1, 'name': 'Groupe Direction Générale', 'is_active': True},
          {'id': 2, 'name': 'RH', 'is_active': True}]
USERS = [{'username': 'william', 'is_ceo': True, 'display_name': 'William',
          'group_ids': [1], 'managed_group_ids': [1]},
         {'username': 'employee', 'display_name': 'Alice', 'group_ids': [2], 'managed_group_ids': []}]


class DirectoryApiTests(unittest.TestCase):
    def setUp(self):
        self.actor = {'username': 'employee', 'role_tags': ['Employe']}
        self.cursor = MagicMock()
        # These scope tests model an older installation without the optional register.
        self.cursor.fetchone.return_value = None
        self.connection = MagicMock()
        self.connection.cursor.return_value.__enter__.return_value = self.cursor
        self.connect = MagicMock(return_value=self.connection)
        def authenticated():
            if self.actor is None:
                raise HTTPException(401, 'Authentication required')
            return self.actor
        app = FastAPI()
        register_team_directory(app, get_db_connection=self.connect, get_current_user=authenticated)
        self.client = TestClient(app)
        self.addCleanup(self.client.close)

    def test_both_routes_require_authentication(self):
        self.actor = None
        for path in ['/organization-directory', '/my-team']:
            self.assertEqual(self.client.get(path).status_code, 401)
        self.connect.assert_not_called()

    def test_all_roles_can_read_sanitized_organization(self):
        for role in ['Employe', 'Responsable', 'Admin']:
            self.actor['role_tags'] = [role]
            self.cursor.fetchall.side_effect = [GROUPS, [{'username': 'a', 'email': 'a@example.test'}], []]
            response = self.client.get('/organization-directory')
            self.assertEqual(response.status_code, 200)
            self.assertNotIn('@', str(response.json()))
            self.assertNotIn('email', response.json()['users'][0])
        self.connection.commit.assert_not_called()

    def test_own_team_scope_cannot_be_overridden_by_query(self):
        self.cursor.fetchall.side_effect = [[GROUPS[1]], [
            {'username': 'lead', 'email': 'lead@example.test', 'group_id': 2, 'relation': 'manager'},
            {'username': 'employee', 'email': 'a@example.test', 'group_id': 2, 'relation': 'member'}]]
        response = self.client.get('/my-team?username=outsider&group_id=1')
        self.assertEqual(response.status_code, 200)
        calls = self.cursor.execute.call_args_list
        self.assertEqual(calls[0].args[1], ('employee',))
        self.assertEqual(calls[1].args[1], (2,))
        self.assertEqual(response.json()['members'][0]['managed_group_ids'], [2])
        self.connection.close.assert_called_once()
        self.connection.commit.assert_not_called()

    def test_no_active_department_returns_empty(self):
        self.cursor.fetchall.return_value = []
        self.assertEqual(self.client.get('/my-team').json(), {'groups': [], 'members': []})
        self.assertEqual(self.cursor.execute.call_count, 1)

    def test_database_failure_closes_connection(self):
        self.cursor.execute.side_effect = RuntimeError('database failure')
        with self.assertRaises(RuntimeError):
            self.client.get('/my-team')
        self.connection.close.assert_called_once()


class DirectoryModelTests(unittest.TestCase):
    def test_prefix_fallback_and_full_name(self):
        self.assertEqual(person_label({'email': 'alice@example.test'}), 'alice')
        self.assertEqual(person_label({'full_name': 'Alice Dupont'}), 'Alice Dupont')
        self.assertEqual(person_label({'username': 'private-login', 'first_name': 'Alice',
                                      'last_name': 'Dupont'}), 'Nom non renseigné')

    def test_leader_listed_once_as_member(self):
        model = build_organization(USERS, GROUPS)
        self.assertEqual(model['departments'][0]['members'], ['William'])
        self.assertEqual(model['departments'][0]['leaders'], ['William'])

    def test_inactive_and_conflicting_assignments_are_retained_as_warnings(self):
        users = USERS + [{'username': 'orphan', 'group_ids': [3]},
                         {'username': 'legacy', 'group_ids': [1, 2], 'managed_group_ids': [1]}]
        model = build_organization(users, GROUPS + [{'id': 3, 'name': 'Old', 'is_active': False}])
        self.assertEqual(len(model['departments']), 2)
        self.assertEqual(model['unassigned'][0]['Situation'], 'Département inactif')
        self.assertGreaterEqual(len(model['issues']), 2)
        self.assertEqual(users[-1]['group_ids'], [1, 2])


    def test_private_fields_discarded_and_memberships_deduplicated(self):
        rows = [{'username': 'a', 'email': 'a@example.test', 'password': 'private'}]
        links = [{'username': 'a', 'group_id': 2, 'relation': 'member'}] * 2
        projected = directory_people(rows, links)
        self.assertEqual(projected[0]['group_ids'], [2])
        self.assertNotIn('password', projected[0])
        members = team_members([dict(rows[0], group_id=2, relation='member')] * 2)
        self.assertEqual(len(members), 1)


if __name__ == '__main__':
    unittest.main()
