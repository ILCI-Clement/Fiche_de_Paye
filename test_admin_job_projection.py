"""Administration job labels with synthetic rows; never connects to a database."""
import copy
import unittest
from unittest.mock import MagicMock, patch

import pymysql

import test_backend_helpers  # isolated import defaults; no connection is opened
from backend import main


def account(username, tags):
    return {'username': username, 'email': f'{username}@example.test',
            'full_name': 'Same synthetic name', 'role': tags[0], 'role_tags': tags}


class AdminJobProjectionTests(unittest.TestCase):
    def setUp(self):
        self.users = [account('opaque-one', ['Employe']), account('opaque-two', ['Employe'])]
        self.rows = [
            {'linked_username': 'opaque-one', 'job_title': '  Developer  '},
            {'linked_username': 'opaque-two', 'job_title': 'Designer'},
            {'linked_username': None, 'full_name': 'Same synthetic name', 'job_title': 'Unlinked'},
        ]
        self.db = MagicMock()
        self.cursor = self.db.cursor.return_value.__enter__.return_value
        self.cursor.fetchone.return_value = {'table': 'personnel_register'}
        self.cursor.fetchall.return_value = self.rows
        self.patches = [
            patch.object(main, 'get_db_connection', return_value=self.db),
            patch.object(main, 'all_users', side_effect=lambda _cursor: copy.deepcopy(self.users)),
            patch.object(main, 'user_can_manage', return_value=True),
        ]
        for item in self.patches:
            item.start()
            self.addCleanup(item.stop)

    def list_as(self, tags, username='opaque-one'):
        return main.list_users(account(username, tags))['users']

    def test_admin_sees_only_explicitly_linked_titles_without_changing_identity(self):
        result = self.list_as(['Admin'])
        self.assertEqual([(user['id'], user['job_title']) for user in result],
                         [('opaque-one', 'Developer'), ('opaque-two', 'Designer')])
        self.assertEqual([user['username'] for user in result], ['opaque-one', 'opaque-two'])
        statements = [call.args[0] for call in self.cursor.execute.call_args_list]
        self.assertIn('SELECT linked_username,job_title FROM personnel_register', statements[-1])
        self.assertNotIn('full_name', statements[-1])
        self.assertNotIn('social_security', repr(result))

    def test_superadmin_inherits_title_projection(self):
        self.assertEqual(self.list_as(['SuperAdmin'])[0]['job_title'], 'Developer')

    def test_manager_and_employee_never_read_or_receive_register_job_titles(self):
        for tags in [['Responsable'], ['Employe']]:
            with self.subTest(tags=tags):
                self.cursor.reset_mock()
                result = self.list_as(tags)
                self.assertTrue(result)
                self.assertTrue(all('job_title' not in user for user in result))
                self.cursor.execute.assert_not_called()

    def test_shared_public_user_does_not_include_job_title(self):
        user = {**self.users[0], 'job_title': 'Not public'}
        self.assertNotIn('job_title', main.public_user(user))

    def test_missing_optional_register_returns_null_titles(self):
        self.cursor.fetchone.return_value = None
        self.assertEqual([user['job_title'] for user in self.list_as(['Admin'])], [None, None])
        self.cursor.fetchall.assert_not_called()

    def test_unlinked_duplicate_name_and_missing_title_do_not_match_accounts(self):
        self.cursor.fetchall.return_value = [self.rows[2],
            {'linked_username': 'opaque-one', 'job_title': '  '}]
        self.assertEqual([user['job_title'] for user in self.list_as(['Admin'])], [None, None])

    def test_table_removal_race_is_compatible_but_other_database_errors_propagate(self):
        def execute(sql, _params=None):
            if sql.startswith('SELECT linked_username,job_title'):
                raise pymysql.err.ProgrammingError(1146, 'Missing synthetic table')
        self.cursor.execute.side_effect = execute
        self.assertEqual([user['job_title'] for user in self.list_as(['Admin'])], [None, None])
        self.cursor.execute.side_effect = pymysql.err.ProgrammingError(1064, 'Synthetic syntax failure')
        with self.assertRaises(pymysql.err.ProgrammingError):
            self.list_as(['Admin'])

    def test_direct_nonadmin_helper_has_no_register_access(self):
        self.assertEqual(main.admin_job_titles(self.cursor, account('manager', ['Responsable']), self.users), {})
        self.cursor.execute.assert_not_called()


if __name__ == '__main__':
    unittest.main()
