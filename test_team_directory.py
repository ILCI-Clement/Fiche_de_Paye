"""Read-only directory authorization and UI regression, with no production writes."""
import unittest
from unittest.mock import MagicMock, patch
from fastapi import FastAPI, HTTPException
from fastapi.testclient import TestClient
from streamlit.testing.v1 import AppTest
from backend.team_directory import register_team_directory, directory_people, team_members
from organization_model import build_organization, person_label
from organization_cards import organization_html
from team_ui import department_rows

GROUPS = [{'id': 1, 'name': 'Groupe Direction Générale', 'is_active': True},
          {'id': 2, 'name': 'RH', 'is_active': True}]
USERS = [{'username': 'william', 'is_ceo': True, 'display_name': 'William',
          'group_ids': [1], 'managed_group_ids': [1]},
         {'username': 'employee', 'display_name': 'Alice', 'group_ids': [2], 'managed_group_ids': []}]


class DirectoryApiTests(unittest.TestCase):
    def setUp(self):
        self.actor = {'username': 'employee', 'role_tags': ['Employe']}
        self.cursor = MagicMock()
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
        self.assertEqual(person_label({'first_name': 'Alice', 'last_name': 'Dupont'}), 'Alice Dupont')

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

    def test_html_escapes_names_and_uses_responsive_static_cards(self):
        model = build_organization(USERS, GROUPS)
        model['departments'][0]['members'] = ['<script>alert(1)</script>']
        html = organization_html(model)
        self.assertNotIn('<script>', html)
        self.assertIn('&lt;script&gt;', html)
        self.assertIn('@container (max-width:480px)', html)
        self.assertNotIn('<svg', html)

    def test_private_fields_discarded_and_memberships_deduplicated(self):
        rows = [{'username': 'a', 'email': 'a@example.test', 'password': 'private'}]
        links = [{'username': 'a', 'group_id': 2, 'relation': 'member'}] * 2
        projected = directory_people(rows, links)
        self.assertEqual(projected[0]['group_ids'], [2])
        self.assertNotIn('password', projected[0])
        members = team_members([dict(rows[0], group_id=2, relation='member')] * 2)
        self.assertEqual(len(members), 1)

    def test_leader_department_shown_in_team_table(self):
        rows = department_rows({'groups': GROUPS, 'members': [
            {'name': 'Lead', 'group_ids': [], 'managed_group_ids': [2]}]})
        self.assertEqual(rows[0]['Département'], 'RH')


class PeoplePageTests(unittest.TestCase):
    def page(self, role, *, organization_failure=False):
        responses = {
            'organization-directory': None if organization_failure else {'groups': GROUPS, 'users': USERS},
            'my-team': {'groups': [GROUPS[1]], 'members': [
                {'name': 'Alice', 'group_ids': [2], 'managed_group_ids': []}]},
            'list-users': {'users': [{'username': 'employee', 'email': 'a@example.test', 'role': 'Employe', 'group_ids': [2]}]},
            'groups': {'groups': GROUPS},
        }
        app = AppTest.from_file('pages/People.py')
        app.session_state['user'] = {'username': 'employee', 'role_tags': [role]}
        with patch('team_api.fetch_directory', side_effect=lambda path: responses[path]) as fetch, \
                patch('streamlit.page_link'):
            app.run()
        self.assertFalse(app.exception)
        return app, [call.args[0] for call in fetch.call_args_list]

    def test_employee_sees_own_list_without_personnel_request(self):
        app, paths = self.page('Employe')
        self.assertEqual(paths, ['organization-directory', 'my-team'])
        self.assertEqual(app.dataframe[0].value.iloc[0]['Nom'], 'Alice')
        self.assertFalse(app.expander[0].proto.expanded)
        self.assertEqual(len(app.segmented_control), 0)

    def test_admin_and_manager_retain_existing_scope(self):
        for role in ['Admin', 'Responsable']:
            app, paths = self.page(role)
            self.assertEqual(paths, ['organization-directory', 'list-users', 'groups'])
            self.assertIn('Utilisateur', app.dataframe[0].value.columns)

    def test_directory_unavailable_does_not_fabricate_chart(self):
        app, paths = self.page('Employe', organization_failure=True)
        self.assertEqual(len(app.expander), 0)
        self.assertEqual(len(app.dataframe), 1)


if __name__ == '__main__':
    unittest.main()
