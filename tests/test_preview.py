"""Verify the preview never requires production secrets or network access."""

from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import TestCase
from unittest.mock import patch

from streamlit.testing.v1 import AppTest
from app_core import create_form, create_person, ensure_workspace
from preview_storage import empty_workspace, load_preview_workspace, save_preview_workspace

ROOT = Path(__file__).resolve().parents[1]


class PreviewTests(TestCase):
    def test_local_save_survives_new_load(self):
        with TemporaryDirectory() as folder:
            database = Path(folder) / 'preview.sqlite3'
            workspace = load_preview_workspace(database)
            self.assertFalse(database.exists())
            workspace['people'].append(create_person('employee', last_name='Modification locale'))
            save_preview_workspace(workspace, database)
            self.assertEqual(load_preview_workspace(database)['people'][0]['last_name'], 'Modification locale')

    def test_old_demonstration_people_are_removed_without_losing_user_data(self):
        with TemporaryDirectory() as folder:
            database = Path(folder) / 'preview.sqlite3'
            demo = create_person('employee', first_name='Camille', last_name='Exemple', supervisor_email='responsable@example.invalid')
            user_person = create_person('employee', first_name='Marie', last_name='Durand')
            workspace = ensure_workspace({
                'people': [demo, user_person],
                'forms': [create_form(2026, 9, demo), create_form(2026, 9, user_person)],
            })
            save_preview_workspace(workspace, database)
            cleaned = load_preview_workspace(database)
            self.assertEqual([person['last_name'] for person in cleaned['people']], ['Durand'])
            self.assertEqual([form['last_name'] for form in cleaned['forms']], ['Durand'])
            self.assertEqual(load_preview_workspace(database), cleaned)

    def test_home_to_forms_without_any_secrets_or_network(self):
        app = AppTest.from_file(str(ROOT / 'local_preview.py'), default_timeout=20)
        with patch('requests.get', side_effect=AssertionError('No production reads')), patch('requests.post', side_effect=AssertionError('No production writes')), patch('preview_storage.load_preview_workspace', return_value=empty_workspace()):
            app.run()
            self.assertFalse(list(app.exception))
            next(button for button in app.button if button.label == 'Choisir une personne').click().run()
            self.assertFalse(list(app.exception))
            self.assertEqual(len(app.tabs), 5)
            self.assertFalse(app.session_state['workspace_local-preview']['people'])
            self.assertFalse(app.session_state['workspace_local-preview']['forms'])
            app.session_state['workspace_local-preview']['forms'].append(create_form(2026, 10, person_type='employee'))
            app.switch_page('pages/Fiches.py').run()
            with patch('preview_storage.save_preview_workspace') as save:
                next(button for button in app.button if button.label == 'Sauvegarder les données de test sur cet ordinateur').click().run()
                self.assertFalse(list(app.exception))
                save.assert_called_once()
