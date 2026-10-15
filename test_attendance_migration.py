"""New attendance adapter compatibility without database writes or provider calls."""
import copy
import io
import json
import unittest
import zipfile
from unittest.mock import MagicMock, Mock, patch
from fastapi import FastAPI
from fastapi.testclient import TestClient
from backend.attendance_api import register_attendance_routes


class AttendanceMigrationTests(unittest.TestCase):
    def setUp(self):
        self.actor = {'username': 'actor', 'role_tags': ['Admin']}
        self.config = {'mois': 9, 'annee': 2026, 'legacy': {'preserve': True}, 'employes_data': [
            {'id': 'one', 'nom': 'Unit fixture', 'ddc': '2026-01-01', 'fdc': '2026-12-31',
             'responsable': 'Manager', 'legacy_unknown': 'preserve'},
            {'id': 'two', 'nom': 'Second unit fixture', 'ddc': '2026-01-01', 'fdc': '2026-12-31', 'responsable': 'Manager'}]}
        db = MagicMock()
        cursor = db.cursor.return_value.__enter__.return_value
        cursor.fetchone.side_effect = lambda: {'form_content': json.dumps(self.config)}
        self.save = Mock()
        self.access = Mock(return_value=True)
        app = FastAPI()
        register_attendance_routes(app, get_db_connection=lambda: db, get_current_user=lambda: self.actor,
            load_user=lambda *_: {'username': 'fixture'}, can_manage=self.access, save_config=self.save)
        self.client = TestClient(app)
        self.addCleanup(self.client.close)

    def export(self, ids, format='original'):
        return self.client.post('/attendance/export', json={'ids': ids, 'month': 9, 'year': 2026, 'format': format})

    def generator(self, employee, month, year, ordinal=None):
        return {'data': employee['id'].encode(), 'filename': f"{employee['id']}.xlsx", 'mime': 'application/mock'}

    def test_export_contains_only_selected_sheet(self):
        with patch('backend.attendance_api.generate_fiche_file', side_effect=self.generator):
            self.assertEqual(self.export(['two']).content, b'two')

    def test_batch_zip_and_duplicate_ids(self):
        with patch('backend.attendance_api.generate_fiche_file', side_effect=self.generator):
            response = self.export(['one', 'two', 'one'])
        with zipfile.ZipFile(io.BytesIO(response.content)) as archive:
            self.assertEqual(archive.namelist(), ['one.xlsx', 'two.xlsx'])

    def test_pdf_uses_existing_converter(self):
        with patch('backend.attendance_api.generate_fiche_file', side_effect=self.generator), patch('backend.esign.convert_fiche_to_pdf', return_value=b'%PDF-mock') as convert:
            response = self.export(['one'], 'pdf')
        self.assertEqual(response.content, b'%PDF-mock')
        convert.assert_called_once_with(b'one', 'one.xlsx')

    def test_export_rejects_other_contract_month(self):
        self.config['employes_data'][0].update(contract_id=1, sheet_month=10, sheet_year=2026)
        self.assertEqual(self.export(['one']).status_code, 400)

    def test_employee_cannot_export_or_edit(self):
        self.actor['role_tags'] = ['Employe']
        self.assertEqual(self.export(['one']).status_code, 403)
        response = self.client.post('/attendance/calendar/one', json={'month': 9, 'year': 2026, 'cells': [{'date': '2026-09-01', 'period': 'matin'}], 'status': 'Congé payé'})
        self.assertEqual(response.status_code, 403)
        self.save.assert_not_called()

    def test_manager_cannot_export_outside_team(self):
        self.actor['role_tags'] = ['Responsable']
        self.access.return_value = False
        self.assertEqual(self.export(['one']).status_code, 403)

    def test_calendar_preserves_unknown_legacy_fields(self):
        original = copy.deepcopy(self.config)
        response = self.client.post('/attendance/calendar/one', json={'month': 9, 'year': 2026, 'cells': [{'date': '2026-09-01', 'period': 'matin'}], 'status': 'Congé payé'})
        self.assertEqual(response.status_code, 200)
        saved = self.save.call_args.args[1]
        self.assertEqual(saved['legacy'], original['legacy'])
        self.assertEqual(saved['employes_data'][0]['legacy_unknown'], 'preserve')
        self.assertEqual(saved['employes_data'][1], original['employes_data'][1])

    def test_calendar_rejects_other_month_without_save(self):
        response = self.client.post('/attendance/calendar/one', json={'month': 9, 'year': 2026, 'cells': [{'date': '2026-10-01', 'period': 'matin'}], 'status': 'Congé payé'})
        self.assertEqual(response.status_code, 400)
        self.save.assert_not_called()


if __name__ == '__main__':
    unittest.main()
