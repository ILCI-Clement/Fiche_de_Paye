"""Streamlit smoke tests with synthetic accounts; no remote calls."""

from pathlib import Path
from unittest import TestCase
from unittest.mock import patch
from streamlit.testing.v1 import AppTest
from app_core import ensure_workspace
from test_attendance import sample_form

ROOT = Path(__file__).resolve().parents[1]


class InterfaceTests(TestCase):
    def create_app(self):
        app = AppTest.from_file(str(ROOT / 'pages' / 'Fiches.py'), default_timeout=20)
        app.secrets['URL_PRESENCE'] = 'http://unused.invalid'
        app.secrets['PRESENCE_TOKEN'] = 'test-only'
        app.session_state['user'] = {'name': 'tester', 'role': 'Responsable'}
        app.session_state['workspace_tester'] = ensure_workspace({'forms': [sample_form()]})
        return app

    def test_form_render_and_completion(self):
        app = self.create_app()
        with patch('requests.get', side_effect=AssertionError('No live network')):
            app.run()
        self.assertFalse(list(app.exception))
        self.assertFalse(any(item.label == 'Jour à modifier' for item in app.selectbox))
        day_buttons = [button for button in app.button if str(button.key).startswith('select_day_')]
        self.assertEqual(len(day_buttons), 30)
        day_buttons[1].click().run()
        self.assertFalse(list(app.exception))
        self.assertEqual(app.session_state[f"selected_day_{app.session_state['workspace_tester']['forms'][0]['id']}"], 2)
        complete = next(button for button in app.button if button.label == 'Marquer terminé')
        complete.click().run()
        self.assertFalse(list(app.exception))
        self.assertEqual(app.session_state['workspace_tester']['forms'][0]['state'], 'complete')
        prepare = next(button for button in app.button if button.label == 'Préparer les exports')
        prepare.click().run()
        self.assertFalse(list(app.exception))
        self.assertEqual(app.session_state['workspace_tester']['forms'][0]['state'], 'complete')
        self.assertTrue(app.session_state['exports_tester']['files'])

    def test_employee_cannot_open_management_page(self):
        app = self.create_app()
        app.session_state['user'] = {'name': 'tester', 'role': 'Employe'}
        app.run()
        self.assertFalse(list(app.exception))
        self.assertFalse(list(app.tabs))
        self.assertTrue(list(app.error))

    def test_loading_failure_never_caches_empty_data(self):
        app = AppTest.from_file(str(ROOT / 'pages' / 'Fiches.py'), default_timeout=20)
        app.secrets['URL_PRESENCE'] = 'http://unused.invalid'
        app.secrets['PRESENCE_TOKEN'] = 'test-only'
        app.session_state['user'] = {'name': 'tester', 'role': 'Responsable'}
        import requests
        with patch('requests.get', side_effect=requests.ConnectionError('offline')):
            app.run()
        self.assertFalse(list(app.exception))
        self.assertNotIn('workspace_tester', app.session_state)
        self.assertTrue(list(app.error))

    def test_archived_edit_stays_archived(self):
        app = self.create_app()
        form = app.session_state['workspace_tester']['forms'][0]
        form.update(state='archived', archived=True)
        app.run()
        next(item for item in app.checkbox if item.label == 'Modifier cette fiche archivée').check().run()
        self.assertFalse(list(app.exception))
        app.text_input(key=f"archived_{form['id']}_last").input('Dupont').run()
        self.assertFalse(list(app.exception))
        form = app.session_state['workspace_tester']['forms'][0]
        self.assertEqual(form['last_name'], 'Dupont')
        self.assertEqual(form['state'], 'archived')

    def test_trash_requires_explicit_confirmation(self):
        from app_core import move_to_trash
        app = self.create_app()
        workspace = app.session_state['workspace_tester']
        move_to_trash(workspace, workspace['forms'][0]['id'])
        app.run()
        next(item for item in app.button if item.label == 'Supprimer définitivement').click().run()
        self.assertFalse(list(app.exception))
        self.assertEqual(len(app.session_state['workspace_tester']['trash']), 1)
        next(item for item in app.button if item.label == 'Confirmer la suppression définitive').click().run()
        self.assertFalse(list(app.exception))
        self.assertEqual(len(app.session_state['workspace_tester']['trash']), 0)
