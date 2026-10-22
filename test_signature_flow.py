import unittest
from unittest.mock import MagicMock
from fastapi import HTTPException
from backend.signature_flow import resolve_manager, signing_people, attendance_signature_fields
from unittest.mock import patch


class SignatureFlowTests(unittest.TestCase):
    def test_pdf_and_setup_links_are_not_signature_buttons(self):
        from backend.esign import sandbox_signing_links, sandbox_signature_actions
        reference = 'unit-document'
        employee = f'https://esign.clawshow.ai/esign/{reference}?token=employee-unit'
        manager = f'https://esign.clawshow.ai/esign/{reference}?token=manager-unit'
        reply = {'document_id': reference, 'sign_url': employee,
                 'pdf_url': f'https://esign.clawshow.ai/esign/{reference}/document.pdf?token=pdf-unit',
                 'setup_url': f'https://esign.clawshow.ai/esign/{reference}/setup',
                 'signers': [{'email':'employee@example.com','url': employee},
                             {'email':'manager@example.com','url': manager}]}
        self.assertEqual(sandbox_signing_links(reply), [employee, manager])
        self.assertEqual([a['party'] for a in sandbox_signature_actions(reply, 'employee@example.com', 'manager@example.com')], ['employee', 'manager'])

    def test_unknown_link_owner_does_not_invent_signer_role(self):
        from backend.esign import sandbox_signature_actions
        reply = {'document_id':'unit', 'url':'https://esign.clawshow.ai/esign/unit?token=unit'}
        self.assertEqual(sandbox_signature_actions(reply, 'e@example.com', 'm@example.com')[0]['party'], 'unknown')

    def test_sandbox_links_only_use_provider_https_origins(self):
        from backend.esign import sandbox_signing_links
        reply = {'sign_url': 'https://esign.clawshow.ai/sign/test', 'signers': [
            {'url': 'https://esign.clawshow.ai/sign/test'},
            {'url': 'javascript:alert(1)'}, {'url': 'https://clawshow.ai.evil.example/sign'},
            {'url': 'https://user:pass@esign.clawshow.ai/sign'}],
            'download_url': 'https://example.com/file.pdf'}
        self.assertEqual(sandbox_signing_links(reply), ['https://esign.clawshow.ai/sign/test'])

    def test_email_names_are_sanitized_for_both_signers(self):
        signers=signing_people('christophe@example.com','christophe@example.com','william@example.com','william@example.com')
        self.assertEqual([s['name'] for s in signers],['christophe','william'])
        self.assertEqual(signers[1]['email'],'william@example.com')

    def test_real_names_are_preserved(self):
        signers=signing_people('HUANG Christophe','christophe@example.com','William TANG','william@example.com')
        self.assertEqual(signers[1]['name'],'William TANG')

    def test_blank_name_uses_mailbox_label(self):
        self.assertEqual(signing_people('', 'christophe@example.com', '', 'william@example.com')[1]['name'],'william')

    def resolve(self, manager_id='boss', email='boss@example.com', access=None):
        db=MagicMock(); cursor=db.cursor.return_value.__enter__.return_value
        cursor.fetchall.return_value=[{'username':'employee'}]
        manager={'username':'boss','email':email,'role_tags':['Admin']}
        employee={'username':'employee','manager_username':manager_id}
        return resolve_manager('employee@example.com', {'username':'manager'}, lambda:db,
            lambda *args:manager, access or (lambda *args,**kwargs:employee))

    def test_signer_order_and_roles_follow_provider_example(self):
        signers=signing_people('Employee','employee@example.com','Boss','boss@example.com')
        self.assertEqual([(s['order'],s['role']) for s in signers],[(1,'SIGNER'),(2,'SIGNER')])

    def test_manager_is_resolved_from_account(self):
        self.assertEqual(self.resolve()['email'],'boss@example.com')

    def test_no_manager_stops_request(self):
        with self.assertRaises(HTTPException) as error:
            self.resolve(manager_id=None)
        self.assertEqual(error.exception.status_code,400)

    def test_same_email_is_rejected(self):
        with self.assertRaises(HTTPException):
            self.resolve(email='employee@example.com')

    def test_outside_team_is_rejected_before_signature_request(self):
        def forbidden(*args,**kwargs): raise HTTPException(403,'outside team')
        with self.assertRaises(HTTPException) as error:
            self.resolve(access=forbidden)
        self.assertEqual(error.exception.status_code,403)

    def test_provider_request_contains_both_signers(self):
        import test_backend_helpers
        from backend import main
        import json
        reply=MagicMock(); reply.__enter__.return_value.read.return_value=b'{"success":true,"document_id":"mock"}'
        with patch.dict('os.environ',{'CLAWSHOW_ESIGN_API_KEY':'test-only'}), patch('backend.esign.place_signature_fields',return_value=attendance_signature_fields()), patch('backend.esign.urllib.request.urlopen',return_value=reply) as request:
            main.submit_to_clawshow(document_url='https://example.com/file.pdf',employee_name='Employee',employee_email='employee@example.com',manager_name='Boss',manager_email='boss@example.com',month=9,year=2026,pdf_bytes=b'%PDF-test')
            payload=json.loads(request.call_args.args[0].data)
            self.assertEqual(len(payload['signers']),2)
            self.assertEqual(payload['signers'][1]['email'],'boss@example.com')
            self.assertEqual(payload['signature_level'],'aes')
            self.assertEqual(payload['signature_fields'][1]['party'],1)
            self.assertEqual(payload['signature_fields'][1]['type'],'SIGNATURE')
            self.assertEqual(payload['schoolInfo']['representative']['name'],'Boss')
            self.assertNotIn('school_full',payload)

    def test_anchors_follow_employee_and_intern_templates(self):
        self.assertEqual(attendance_signature_fields()[0]['anchor'],'Signature du salarié :')
        self.assertEqual(attendance_signature_fields(intern=True)[0]['anchor'],'Signature du stagiaire :')
        self.assertEqual([f['party'] for f in attendance_signature_fields()],[0,1])
