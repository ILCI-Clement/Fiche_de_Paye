"""Smoke tests for the production login and role boundaries."""

from pathlib import Path
from unittest import TestCase
from unittest.mock import Mock, patch

from streamlit.testing.v1 import AppTest


ROOT = Path(__file__).resolve().parents[1]


class AccessTests(TestCase):
    def production_app(self, user=None):
        app = AppTest.from_file(str(ROOT / 'online_test.py'), default_timeout=20)
        app.secrets['URL_PRESENCE'] = 'http://unused.invalid'
        app.secrets['PRESENCE_TOKEN'] = 'test-only'
        if user is not None:
            app.session_state['user'] = user
        return app

    def test_anonymous_user_sees_login_only(self):
        app = self.production_app()
        app.run()
        self.assertFalse(list(app.exception))
        self.assertEqual([title.value for title in app.title], ['Page de Connexion'])
        self.assertTrue(any(item.label == "Nom d'utilisateur" for item in app.text_input))
        self.assertTrue(any(item.label == 'Mot de passe' for item in app.text_input))

    def test_employee_is_routed_to_profile(self):
        app = self.production_app({'name': 'employee-test', 'email': 'employee@example.invalid', 'role': 'Employe'})
        app.run()
        self.assertFalse(list(app.exception))
        self.assertEqual([title.value for title in app.title], ['Mes Infos Personnelles'])

    def test_administration_keeps_three_account_roles(self):
        app = AppTest.from_file(str(ROOT / 'pages' / 'Admin.py'), default_timeout=20)
        app.secrets['URL_PRESENCE'] = 'http://unused.invalid'
        app.secrets['PRESENCE_TOKEN'] = 'test-only'
        app.session_state['user'] = {'name': 'admin-test', 'email': 'admin@example.invalid', 'role': 'Admin'}
        response = Mock(status_code=200)
        response.json.return_value = {'users': []}
        with patch('requests.get', return_value=response):
            app.run()
        self.assertFalse(list(app.exception))
        role = next(item for item in app.radio if item.label == "Role de l'utilisateur")
        self.assertEqual(role.options, ['Admin', 'Responsable', 'Employe'])
