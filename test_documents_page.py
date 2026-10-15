"""Documents UI regression checks with mocked API responses only."""
import unittest
from unittest.mock import patch
from streamlit.testing.v1 import AppTest


class Response:
    status_code = 200
    def __init__(self, data): self.data = data
    def json(self): return self.data


def receipt(identifier, month, *, archived=False, category='transport'):
    return dict(id=identifier, receipt_month=month, receipt_year=2026,
                original_filename=f'file-{identifier}.pdf', employee_username='admin',
                uploaded_by='admin', created_at='2026-10-14T10:00:00',
                archived=archived, document_type=category, mime_type='application/pdf')


class DocumentsPageTests(unittest.TestCase):
    def setUp(self):
        self.receipts = [receipt(1, 10), receipt(2, 9), receipt(3, 10, archived=True), receipt(4, 10, category='autre')]
        def request(method, url, **kwargs):
            return Response({'users': ['admin']} if url.endswith('/assignees') else {'receipts': self.receipts})
        self.patches = [
            patch('api_client.api_url', return_value='https://api.test'),
            patch('api_client.authenticated_headers', return_value={}),
            patch('api_client.safe_api_request', side_effect=request),
            patch('document_ui.safe_api_request', side_effect=request),
        ]
        for mock in self.patches: mock.start()
        self.addCleanup(lambda: [mock.stop() for mock in reversed(self.patches)])
        self.app = AppTest.from_file('pages/TransportReceipts.py')
        self.app.session_state['user'] = {'name':'admin', 'role':'Admin', 'role_tags':['Admin']}
        self.app.session_state['period_transport'] = (10,2026)
        self.app.run()
        self.assertFalse(self.app.exception)

    def widget(self, kind, key):
        return next(w for w in getattr(self.app,kind) if w.key==key)

    def selected_ids(self):
        return {w.key for w in self.app.checkbox if w.key.startswith('receipt_selected_')}

    def test_month_selectors_are_synchronized_and_filter_bound_month(self):
        self.assertEqual(self.selected_ids(), {'receipt_selected_1','receipt_selected_4'})
        self.widget('selectbox','period_transport_bottom_month').select(9).run()
        self.assertFalse(self.app.exception)
        self.assertEqual(self.widget('selectbox','period_transport_top_month').value,9)
        self.assertEqual(self.selected_ids(),{'receipt_selected_2'})
        self.widget('selectbox','period_transport_top_month').select(10).run()
        self.assertEqual(self.widget('selectbox','period_transport_bottom_month').value,10)

    def test_archives_are_separate(self):
        self.widget('segmented_control','receipt_folder').set_value('Archives').run()
        self.assertFalse(self.app.exception)
        self.assertEqual(self.selected_ids(),{'receipt_selected_3'})

    def test_select_all_respects_filter_and_clears_on_month_change(self):
        self.widget('selectbox','receipt_type_filter').select('Transport').run()
        self.widget('button','receipt_select_all').click().run()
        self.assertEqual(self.selected_ids(),{'receipt_selected_1'})
        self.assertTrue(self.widget('checkbox','receipt_selected_1').value)
        self.widget('selectbox','period_transport_bottom_month').select(9).run()
        self.assertFalse(self.app.exception)
        self.assertFalse(self.widget('checkbox','receipt_selected_2').value)
        self.assertTrue(self.widget('button','receipt_prepare_zip').disabled)

    def test_zip_uses_selection_and_is_invalidated_after_deselect(self):
        self.widget('checkbox','receipt_selected_1').check().run()
        with patch('document_ui.build_document_archive', return_value=b'zip-test') as build:
            self.widget('button','receipt_prepare_zip').click().run()
        self.assertFalse(self.app.exception)
        self.assertEqual([r['id'] for r in build.call_args.args[0]],[1])
        self.assertEqual(self.app.session_state['receipt_bulk_zip']['data'],b'zip-test')
        self.widget('checkbox','receipt_selected_1').uncheck().run()
        self.assertNotIn('receipt_bulk_zip',self.app.session_state)

    def test_no_documents_still_has_month_selector(self):
        self.widget('selectbox','period_transport_bottom_month').select(8).run()
        self.assertFalse(self.app.exception)
        self.assertEqual(self.selected_ids(),set())
        self.assertEqual(self.widget('selectbox','period_transport_bottom_month').value,8)

    def test_employee_upload_is_bound_to_self_without_assignee_selector(self):
        self.app.session_state['user'] = {'name': 'employee', 'role': 'Employe', 'role_tags': ['Employe']}
        self.app.run()
        self.assertFalse(self.app.exception)
        self.assertFalse(any(w.label == 'Employé concerné' for w in self.app.selectbox))
        self.assertTrue(any('employee' in item.value for item in self.app.caption))
        self.assertTrue(any(w.label.startswith('Envoyer pour') for w in self.app.button))

    def test_manager_can_open_personal_upload_without_assignee_selector(self):
        self.app.session_state['user'] = {'name': 'manager', 'role': 'Responsable', 'role_tags': ['Responsable']}
        self.app.run()
        self.assertFalse(self.app.exception)
        self.assertFalse(any(w.label == 'Employé concerné' for w in self.app.selectbox))
        self.assertTrue(any('manager' in item.value for item in self.app.caption))
        self.assertTrue(any(w.label.startswith('Envoyer pour') for w in self.app.button))

    def test_admin_retains_upload_assignee_selector(self):
        self.assertTrue(any(w.label == 'Employé concerné' for w in self.app.selectbox))


if __name__ == '__main__': unittest.main()
