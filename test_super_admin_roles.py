"""Role/access regressions using synthetic in-memory accounts; no live database."""
import copy
import json
import unittest
from unittest.mock import Mock, patch

import test_backend_helpers  # isolated non-production import defaults
from backend import main
from backend.account_lifecycle import register_account_lifecycle
from backend.role_policy import (active_super_admin_count, can_assign_roles, can_manage_account,
                                 canonical_role_tags, checked_resource_actor, is_admin, is_super_admin)
from fastapi import FastAPI, HTTPException
from fastapi.testclient import TestClient


def account(username, tags, *, active=True, version=0):
    tags = canonical_role_tags(tags)
    return dict(username=username, email=f'{username}@example.invalid', role_tags=tags,
                role=tags[0], is_admin='Admin' in tags, is_active=active, session_version=version,
                login_deleted_at=None, employee_type='salarie', manager_username=None,
                group_ids=[], managed_group_ids=[], created_at=None, password_hash='synthetic-hash',
                contract_start_date=None, contract_end_date=None, is_cdi=False,
                remember_token_hash='synthetic-remember', remember_token_expires=None,
                reset_token='synthetic-reset', reset_token_expires=None)


class Fixture:
    def __init__(self):
        self.users = {'root': account('root', ['SuperAdmin']), 'admin': account('admin', ['Admin']),
                      'worker': account('worker', ['Employe'])}
        self.users['worker']['manager_username'] = 'admin'
        self.actor_name = 'root'
        self.events, self.statements, self.reserved = [], [], set()
        self.fail_audit = False
        self.commits = self.rollbacks = 0

    def actor(self):
        return self.load_user(None, self.actor_name)

    def load_user(self, _cursor, username):
        row = self.users.get(username)
        if not row:
            return None
        return copy.deepcopy(row)

    def connect(self):
        return Connection(self)


class Connection:
    def __init__(self, fixture):
        self.fixture = fixture
        self.before = copy.deepcopy((fixture.users, fixture.events, fixture.reserved))

    def cursor(self): return Cursor(self.fixture)
    def close(self): pass
    def commit(self): self.fixture.commits += 1
    def rollback(self):
        self.fixture.users, self.fixture.events, self.fixture.reserved = copy.deepcopy(self.before)
        self.fixture.rollbacks += 1


class Cursor:
    def __init__(self, fixture): self.fixture, self.rows = fixture, []
    def __enter__(self): return self
    def __exit__(self, *_): pass
    def fetchall(self): return copy.deepcopy(self.rows)
    def fetchone(self): return copy.deepcopy(self.rows[0]) if self.rows else None

    def execute(self, sql, params=()):
        query = ' '.join(sql.split())
        self.fixture.statements.append(query)
        self.rows = []
        if query == 'SELECT username FROM users ORDER BY username FOR UPDATE':
            self.rows = [{'username': key} for key in sorted(self.fixture.users)]
        elif query.startswith('SELECT username,role,role_tags,is_admin,is_active,login_deleted_at'):
            self.rows = list(self.fixture.users.values())
        elif query.startswith('SELECT username FROM account_username_reservations'):
            self.rows = [{'username': params[0]}] if params[0] in self.fixture.reserved else []
        elif query.startswith('INSERT IGNORE INTO account_username_reservations'):
            self.fixture.reserved.add(params[0])
        elif query.startswith('INSERT INTO users '):
            name, email, hashed, _admin, role, employee_type, manager = params
            self.fixture.users[name] = account(name, [role])
            self.fixture.users[name].update(email=email, password_hash=hashed,
                                            employee_type=employee_type, manager_username=manager)
        elif query.startswith('UPDATE users SET role_tags ='):
            self.fixture.users[params[1]]['role_tags'] = json.loads(params[0])
        elif query.startswith('UPDATE users SET role ='):
            role, tags, administrator, employee_type, manager, name = params
            self.fixture.users[name].update(role=role, role_tags=json.loads(tags), is_admin=administrator,
                                            employee_type=employee_type, manager_username=manager)
        elif query.startswith('UPDATE users SET session_version=session_version+1'):
            row = self.fixture.users[params[-1]]
            row.update(session_version=row['session_version'] + 1, remember_token_hash=None,
                       remember_token_expires=None, reset_token=None, reset_token_expires=None)
        elif query.startswith('UPDATE users SET is_active='):
            row = self.fixture.users[params[1]]
            row.update(is_active=bool(params[0]), session_version=row['session_version'] + 1,
                       remember_token_hash=None, remember_token_expires=None, reset_token=None)
        elif query.startswith('UPDATE users SET login_deleted_at='):
            row = self.fixture.users[params[1]]
            row.update(login_deleted_at='synthetic-time', is_active=False,
                       session_version=row['session_version'] + 1, password_hash=params[0],
                       remember_token_hash=None, reset_token=None)
        elif query.startswith('UPDATE users SET email ='):
            self.fixture.users[params[1]]['email'] = params[0]
        elif query.startswith('UPDATE users SET manager_username ='):
            self.fixture.users[params[1]]['manager_username'] = params[0]
        elif query.startswith('INSERT INTO account_role_events') or query.startswith('INSERT INTO account_status_events'):
            if self.fixture.fail_audit:
                raise RuntimeError('Synthetic audit failure')
            self.fixture.events.append((query, params))
        else:
            raise AssertionError('Unexpected synthetic SQL operation')


class RolePolicyTests(unittest.TestCase):
    def test_super_inherits_admin_in_any_supported_trusted_shape(self):
        for user in ({'role_tags': ['SuperAdmin']}, {'role_tags': '["SuperAdmin"]'}, {'role': 'SuperAdmin'}):
            self.assertTrue(is_admin(user))
            self.assertTrue(is_super_admin(user))
        self.assertEqual(canonical_role_tags(['Employe', 'SuperAdmin']), ['SuperAdmin', 'Admin', 'Employe'])

    def test_normal_admin_cannot_delegate_or_manage_super(self):
        normal, root = account('a', ['Admin']), account('s', ['SuperAdmin'])
        self.assertFalse(can_assign_roles(normal, ['SuperAdmin']))
        self.assertFalse(can_assign_roles(normal, ['Admin'], root))
        self.assertFalse(can_manage_account(normal, root))
        self.assertTrue(can_assign_roles(root, ['SuperAdmin']))
        self.assertTrue(can_manage_account(root, normal))

    def test_database_super_tags_are_loaded_with_inherited_admin(self):
        row = account('synthetic-super', ['SuperAdmin'])
        row['role_tags'] = '["SuperAdmin"]'
        row['is_admin'] = False  # Inheritance is authoritative even for an older flag.
        cursor = Mock()
        cursor.fetchone.side_effect = [row, None]
        cursor.fetchall.return_value = []
        loaded = main.load_user(cursor, 'synthetic-super')
        self.assertEqual(loaded['role_tags'], ['SuperAdmin', 'Admin'])
        self.assertEqual(loaded['role'], 'SuperAdmin')
        self.assertTrue(main.public_user(loaded)['is_admin'])

    def test_invalid_role_values_fail_closed(self):
        self.assertEqual(canonical_role_tags(['Unknown', {}, True]), [])
        self.assertFalse(can_assign_roles(account('a', ['Admin']), [{}]))
        self.assertFalse(is_super_admin({'role_tags': '["Admin"]', 'role': 'SuperAdmin'}))
        self.assertFalse(is_admin(None))

    def test_fresh_resource_guard_rejects_inactive_deleted_and_stale(self):
        for changes in ({'is_active': False}, {'login_deleted_at': 'synthetic'}, {'session_version': 2}):
            fixture = Fixture()
            actor = fixture.actor()
            fixture.users['root'].update(changes)
            with self.assertRaises(HTTPException) as caught:
                checked_resource_actor(Cursor(fixture), actor, fixture.load_user, lock=True)
            self.assertEqual(caught.exception.status_code, 401)
            self.assertEqual(fixture.statements[0], 'SELECT username FROM users ORDER BY username FOR UPDATE')

    def test_active_super_count_ignores_disabled_deleted_and_excluded(self):
        fixture = Fixture()
        fixture.users['disabled'] = account('disabled', ['SuperAdmin'], active=False)
        fixture.users['deleted'] = account('deleted', ['SuperAdmin'])
        fixture.users['deleted']['login_deleted_at'] = 'synthetic'
        self.assertEqual(active_super_admin_count(Cursor(fixture)), 1)
        self.assertEqual(active_super_admin_count(Cursor(fixture), exclude_username='root'), 0)


class MainRoleMutationTests(unittest.TestCase):
    def setUp(self):
        self.fixture = Fixture()
        for name, replacement in (
            ('get_db_connection', self.fixture.connect), ('load_user', self.fixture.load_user),
            ('validate_group_ids', lambda *_args, **_kwargs: []),
            ('replace_memberships', lambda *_args: None), ('hash_password', lambda _value: 'synthetic-hash'),
        ):
            context = patch.object(main, name, replacement)
            context.start(); self.addCleanup(context.stop)

    def assert_http(self, status, action):
        with self.assertRaises(HTTPException) as caught: action()
        self.assertEqual(caught.exception.status_code, status)

    def change(self, target, tags, actor='root'):
        return main.update_user_organization(target, {'role_tags': tags}, self.fixture.load_user(None, actor))

    def test_requested_roles_expand_super_and_reject_malformed_payload(self):
        self.assertEqual(main.requested_role_tags({'role_tags': ['SuperAdmin']}), ['SuperAdmin', 'Admin'])
        for value in ('SuperAdmin', [], ['Unknown'], {'SuperAdmin': True}):
            self.assert_http(400, lambda: main.requested_role_tags({'role_tags': value}))

    def test_normal_admin_cannot_self_promote_or_demote_a_super(self):
        self.assert_http(403, lambda: self.change('admin', ['SuperAdmin'], 'admin'))
        self.assert_http(403, lambda: self.change('root', ['Admin'], 'admin'))
        self.assertEqual(self.fixture.users['admin']['role_tags'], ['Admin'])

    def test_super_role_change_revokes_existing_session_and_recovery_credentials(self):
        old_token = main.create_session_token('admin', 0)
        result = self.change('admin', ['SuperAdmin'])
        self.assertEqual(result['user']['role'], 'SuperAdmin')
        row = self.fixture.users['admin']
        self.assertEqual(row['session_version'], 1)
        self.assertIsNone(row['remember_token_hash'])
        self.assertIsNone(row['reset_token'])
        self.assert_http(401, lambda: main.get_current_user('Bearer ' + old_token))
        self.assertTrue(is_super_admin(main.get_current_user('Bearer ' + main.create_session_token('admin', 1))))
        self.assertEqual(len(self.fixture.events), 1)

    def test_roles_unchanged_do_not_revoke_session(self):
        self.change('admin', ['Admin'])
        self.assertEqual(self.fixture.users['admin']['session_version'], 0)
        self.assertFalse(self.fixture.events)

    def test_cannot_remove_last_or_own_super_role(self):
        self.assert_http(409, lambda: self.change('root', ['Admin']))
        self.fixture.users['second'] = account('second', ['SuperAdmin'])
        self.assert_http(400, lambda: self.change('root', ['Admin']))

    def test_audit_failure_rolls_back_role_and_version_change(self):
        self.fixture.fail_audit = True
        with self.assertRaises(RuntimeError): self.change('admin', ['SuperAdmin'])
        self.assertEqual(self.fixture.users['admin']['role_tags'], ['Admin'])
        self.assertEqual(self.fixture.users['admin']['session_version'], 0)
        self.assertEqual(self.fixture.rollbacks, 1)

    def test_stale_or_demoted_admin_cannot_modify_roles(self):
        old_actor = self.fixture.load_user(None, 'admin')
        self.fixture.users['admin']['role_tags'] = ['Employe']
        self.fixture.users['admin']['role'] = 'Employe'
        self.assert_http(403, lambda: main.update_user_organization('worker', {'role_tags': ['Admin']}, old_actor))
        self.fixture.users['admin']['session_version'] = 1
        self.assert_http(401, lambda: main.update_user_organization('worker', {'role_tags': ['Admin']}, old_actor))

    def test_account_creation_cannot_bypass_super_delegation(self):
        data = {'new_username': 'synthetic-new', 'new_mail': 'synthetic-new@example.invalid',
                'new_password': 'synthetic-unit-only', 'role_tags': ['SuperAdmin']}
        self.assert_http(403, lambda: main.create_user(data, self.fixture.load_user(None, 'admin')))
        main.create_user(data, self.fixture.actor())
        self.assertTrue(is_super_admin(self.fixture.users['synthetic-new']))
        self.assertEqual(len(self.fixture.events), 1)

    def test_purged_name_cannot_be_recreated_or_taken_by_profile_rename(self):
        self.fixture.reserved.add('reserved-unit-name')
        data = {'new_username': 'reserved-unit-name', 'new_mail': 'new@example.invalid',
                'new_password': 'synthetic-unit-only', 'role_tags': ['Admin']}
        self.assert_http(409, lambda: main.create_user(data, self.fixture.actor()))
        with patch.object(main, 'verify_password', return_value=True):
            payload = main.ProfileUpdateRequest(current_password='synthetic-current', new_username='reserved-unit-name')
            self.assert_http(409, lambda: main.update_profile('root', payload, self.fixture.actor()))

    def test_ordinary_admin_cannot_change_super_manager(self):
        self.assert_http(403, lambda: main.assign_direct_manager('root', main.DirectManagerRequest(manager_id='admin'),
                                                               self.fixture.load_user(None, 'admin')))

    def test_self_profile_update_keeps_super_role(self):
        with patch.object(main, 'verify_password', return_value=True), patch.object(main, 'migrate_username_in_saved_fiches'):
            payload = main.ProfileUpdateRequest(current_password='synthetic-current', new_email='unit-changed@example.com')
            result = main.update_profile('root', payload, self.fixture.actor())
        self.assertTrue(is_super_admin(self.fixture.users['root']))
        self.assertTrue(result['auth_token'])

    def test_permanent_delete_guards_run_before_purge_service(self):
        payload = main.AccountPurgeRequest(confirmation='worker', reason='synthetic_cleanup')
        with patch.object(main, 'perform_account_purge') as purge:
            self.assert_http(403, lambda: main.delete_user('worker', payload, self.fixture.load_user(None, 'admin')))
            self.assert_http(409, lambda: main.delete_user('worker', payload, self.fixture.actor()))
            self.fixture.users['worker']['is_active'] = False
            wrong = main.AccountPurgeRequest(confirmation='different', reason='synthetic_cleanup')
            self.assert_http(400, lambda: main.delete_user('worker', wrong, self.fixture.actor()))
            purge.assert_not_called()

    def test_permanent_delete_delegates_only_after_all_checks(self):
        self.fixture.users['worker']['is_active'] = False
        payload = main.AccountPurgeRequest(confirmation='worker', reason='synthetic_cleanup')
        with patch.object(main, 'perform_account_purge', return_value={'accounts': 1}) as purge:
            result = main.delete_user('worker', payload, self.fixture.actor())
        self.assertEqual(result, {'status': 'success', 'statistics': {'accounts': 1}})
        self.assertEqual(purge.call_args.kwargs, {'reason': 'synthetic_cleanup'})
        self.assertIn('worker', self.fixture.users)  # Mock service; no actual account is deleted.


class AccountAccessRoleTests(unittest.TestCase):
    def setUp(self):
        self.fixture = Fixture()
        app = FastAPI()
        register_account_lifecycle(app, get_db_connection=self.fixture.connect, get_current_user=self.fixture.actor,
                                   load_user=self.fixture.load_user, public_user=main.public_user,
                                   hash_password=lambda _value: 'synthetic-hash')
        self.client = TestClient(app)

    def test_admin_archives_normal_account_but_cannot_touch_super(self):
        self.fixture.actor_name = 'admin'
        self.assertEqual(self.client.patch('/users/root/activation', json={'is_active': False}).status_code, 403)
        self.assertEqual(self.client.patch('/users/worker/activation', json={'is_active': False}).status_code, 200)
        self.assertFalse(self.fixture.users['worker']['is_active'])
        self.assertEqual(self.fixture.users['worker']['session_version'], 1)

    def test_admin_cannot_delete_login_even_for_inactive_normal_account(self):
        self.fixture.actor_name = 'admin'
        self.fixture.users['worker']['is_active'] = False
        response = self.client.request('DELETE', '/users/worker/login', json={'confirmation': 'worker'})
        self.assertEqual(response.status_code, 403)
        self.assertIsNone(self.fixture.users['worker']['login_deleted_at'])

    def test_super_can_delete_inactive_login_without_removing_identity(self):
        self.fixture.users['worker']['is_active'] = False
        response = self.client.request('DELETE', '/users/worker/login', json={'confirmation': 'worker'})
        self.assertEqual(response.status_code, 200)
        self.assertIsNotNone(self.fixture.users['worker']['login_deleted_at'])
        self.assertIn('worker', self.fixture.users)

    def test_super_cannot_disable_itself(self):
        response = self.client.patch('/users/root/activation', json={'is_active': False})
        self.assertEqual(response.status_code, 400)
        self.assertTrue(self.fixture.users['root']['is_active'])


if __name__ == '__main__':
    unittest.main()
