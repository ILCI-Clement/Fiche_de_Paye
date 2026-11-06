import json
import os
import sqlite3
import tempfile
import unittest
from contextlib import closing
from pathlib import Path
from unittest.mock import Mock, patch

from fastapi import FastAPI
from fastapi.testclient import TestClient
from backend.leave_notifications import LeaveNotifications, ensure_notification_schema, mail_message, review_url, smtp_deliver
from backend.leave_requests import register_leave_routes
from backend.leave_store import ensure_leave_schema


class LeaveNotificationTests(unittest.TestCase):
    def connect(self):
        db = sqlite3.connect(self.path)
        db.row_factory = sqlite3.Row
        return db

    def setUp(self):
        temp = tempfile.TemporaryDirectory(); self.addCleanup(temp.cleanup)
        self.path = Path(temp.name)/'notifications.sqlite3'
        with closing(self.connect()) as db, db:
            db.execute('''CREATE TABLE users(username TEXT PRIMARY KEY,email TEXT,manager_username TEXT,
                role_tags TEXT,is_active INT,login_deleted_at TEXT)''')
            db.executemany('INSERT INTO users VALUES(?,?,?,?,1,NULL)', [
                ('member','member@example.com','boss',json.dumps(['Employe'])),
                ('boss','boss@example.com',None,json.dumps(['Responsable'])),
                ('other','other@example.com',None,json.dumps(['Responsable']))])
        ensure_leave_schema(self.connect);ensure_notification_schema(self.connect)
        self.deliver = Mock(return_value='sent')
        self.service = LeaveNotifications(self.connect,deliver=self.deliver)
        self.actor = {'username':'member','role_tags':['Employe']}
        app = FastAPI()
        register_leave_routes(app,get_db_connection=self.connect,get_current_user=lambda:self.actor,notifications=self.service)
        self.client = TestClient(app,raise_server_exceptions=False);self.addCleanup(self.client.close)
        self.addCleanup(patch.stopall)
        patch.dict(os.environ,{'PRESENCE_APP_URL':'https://rh.example.com/','PRESENCE_SMTP_USER':'rh@example.com'}).start()

    def create(self):
        response = self.client.post('/leave-requests',json=dict(start_date='2026-11-05',end_date='2026-11-05',start_slot='afternoon',end_slot='afternoon',note='PRIVATE NOTE'))
        self.assertEqual(response.status_code,201,response.text)
        return response.json()['id']

    def jobs(self):
        with closing(self.connect()) as db:
            return [dict(row) for row in db.execute('SELECT * FROM leave_notification_outbox ORDER BY event_id')]

    def test_submit_is_durable_and_only_sends_once_without_notes(self):
        identifier = self.create()
        self.assertEqual(self.jobs()[0]['status'],'pending')
        self.deliver.assert_not_called()
        self.assertTrue(self.service.dispatch_one())
        self.assertFalse(self.service.dispatch_one())
        message = self.deliver.call_args.args[0]
        self.assertEqual(message['To'],'boss@example.com')
        body = message.get_body(preferencelist=('plain',)).get_content()
        self.assertIn(f'https://rh.example.com/?leave={identifier}',body)
        self.assertIn('Après-midi',body)
        self.assertNotIn('PRIVATE NOTE',body)
        self.assertEqual(self.jobs()[0]['status'],'sent')
        self.deliver.assert_called_once()

    def test_transfer_resolves_current_manager_and_disabled_or_missing_waits(self):
        self.create()
        with closing(self.connect()) as db, db:
            db.execute("UPDATE users SET manager_username='other' WHERE username='member'")
        self.service.dispatch_one()
        self.assertEqual(self.deliver.call_args.args[0]['To'],'other@example.com')
        self.create()
        with closing(self.connect()) as db, db:
            db.execute("UPDATE users SET is_active=0 WHERE username='other'")
        self.service.dispatch_one()
        self.assertEqual(self.jobs()[-1]['status'],'no_recipient')
        self.assertEqual(self.deliver.call_count,1)

    def test_withdrawn_request_not_sent_and_cancellation_queues_separately(self):
        identifier=self.create()
        self.assertEqual(self.client.post(f'/leave-requests/{identifier}/withdraw').status_code,200)
        self.service.dispatch_one()
        self.assertEqual(self.jobs()[0]['status'],'obsolete')
        self.deliver.assert_not_called()
        identifier=self.create()
        self.actor={'username':'boss','role_tags':['Responsable']}
        self.assertEqual(self.client.post(f'/leave-requests/{identifier}/decision',json={'decision':'approved'}).status_code,200)
        self.service.dispatch_one()
        self.actor={'username':'member','role_tags':['Employe']}
        self.assertEqual(self.client.post(f'/leave-requests/{identifier}/cancellation').status_code,200)
        self.assertEqual(self.client.post(f'/leave-requests/{identifier}/cancellation').status_code,409)
        while self.service.dispatch_one(): pass
        self.assertEqual(self.deliver.call_args.args[0]['Subject'],'Annulation de congé à valider')
        self.assertEqual(len(self.jobs()),4)

    def test_decision_result_goes_to_employee_once_with_comment(self):
        for decision in ('approved','rejected'):
            with self.subTest(decision=decision):
                self.actor={'username':'member','role_tags':['Employe']}
                identifier=self.create()
                self.actor={'username':'boss','role_tags':['Responsable']}
                payload={'decision':decision,'comment':'Décision <validée>\nMerci.'}
                url=f'/leave-requests/{identifier}/decision'
                self.assertEqual(self.client.post(url,json=payload).status_code,200)
                self.assertEqual(self.client.post(url,json=payload).status_code,409)
                while self.service.dispatch_one(): pass
                message=self.deliver.call_args.args[0]
                self.assertEqual(message['To'],'member@example.com')
                self.assertIn('approuvée' if decision=='approved' else 'refusée',message['Subject'])
                body=message.get_body(preferencelist=('plain',)).get_content()
                self.assertIn('Décision <validée>',body)
                self.assertNotIn('PRIVATE NOTE',body)
                html=message.get_body(preferencelist=('html',)).get_content()
                self.assertIn('&lt;validée&gt;',html)
                self.assertIn('Voir les détails',html)
        self.assertEqual(self.deliver.call_count,2)

    def test_cancellation_results_explain_whether_leave_stays_valid(self):
        for decision in ('approved','rejected'):
            with self.subTest(decision=decision):
                self.actor={'username':'member','role_tags':['Employe']}
                identifier=self.create()
                self.actor={'username':'boss','role_tags':['Responsable']}
                self.assertEqual(self.client.post(f'/leave-requests/{identifier}/decision',json={'decision':'approved'}).status_code,200)
                self.actor={'username':'member','role_tags':['Employe']}
                self.assertEqual(self.client.post(f'/leave-requests/{identifier}/cancellation').status_code,200)
                self.actor={'username':'boss','role_tags':['Responsable']}
                url=f'/leave-requests/{identifier}/cancellation-decision'
                self.assertEqual(self.client.post(url,json={'decision':decision,'comment':'Commentaire final'}).status_code,200)
                self.assertEqual(self.client.post(url,json={'decision':decision}).status_code,409)
                while self.service.dispatch_one(): pass
                message=self.deliver.call_args.args[0]
                self.assertEqual(message['To'],'member@example.com')
                self.assertIn('annulé' if decision=='approved' else 'reste approuvé',message['Subject'])
                self.assertIn('Commentaire final',message.get_body(preferencelist=('plain',)).get_content())
        self.assertEqual(self.deliver.call_count,4)

    def test_decision_mail_failure_does_not_undo_approval_or_resend_with_uncertain_result(self):
        identifier=self.create()
        self.actor={'username':'boss','role_tags':['Responsable']}
        self.assertEqual(self.client.post(f'/leave-requests/{identifier}/decision',json={'decision':'approved'}).status_code,200)
        self.deliver.return_value='uncertain'
        while self.service.dispatch_one(): pass
        self.assertEqual(self.jobs()[-1]['status'],'uncertain')
        self.assertEqual(self.client.get(f'/leave-requests/{identifier}').json()['request']['status'],'approved')
        self.deliver.assert_called_once()

    def test_failed_outbox_insert_rolls_back_decision(self):
        identifier=self.create()
        self.actor={'username':'boss','role_tags':['Responsable']}
        with patch.object(self.service,'enqueue',side_effect=RuntimeError('synthetic failure')):
            self.assertEqual(self.client.post(f'/leave-requests/{identifier}/decision',json={'decision':'approved'}).status_code,500)
        self.assertEqual(self.client.get(f'/leave-requests/{identifier}').json()['request']['status'],'pending')
        self.assertEqual(len(self.jobs()),1)

    def test_outbox_failure_rolls_back_request_and_event(self):
        with closing(self.connect()) as db,db:db.execute('DROP TABLE leave_notification_outbox')
        response=self.client.post('/leave-requests',json=dict(start_date='2026-11-05',end_date='2026-11-05'))
        self.assertEqual(response.status_code,500)
        with closing(self.connect()) as db:
            self.assertEqual(db.execute('SELECT COUNT(*) FROM leave_requests').fetchone()[0],0)
            self.assertEqual(db.execute('SELECT COUNT(*) FROM leave_events').fetchone()[0],0)

    def test_uncertain_send_is_not_repeated(self):
        self.create();self.deliver.return_value='uncertain'
        self.service.dispatch_one()
        self.assertEqual(self.jobs()[0]['status'],'uncertain')
        self.assertFalse(self.service.dispatch_one())
        self.deliver.assert_called_once()

    def test_message_escapes_html_and_link_requires_https(self):
        with self.assertRaises(ValueError):review_url('http://example.com',1)
        with self.assertRaises(ValueError):review_url('https://user:password@example.com',1)
        row=dict(id=3,kind='submitted',start_date='2026-11-05',end_date='2026-11-05',start_slot='full',end_slot='full')
        message=mail_message(row,'<script>','https://rh.example.com/?untrusted=1#x','rh@example.com','boss@example.com')
        html=message.get_body(preferencelist=('html',)).get_content()
        self.assertNotIn('<script>',html);self.assertIn('&lt;script&gt;',html)
        self.assertIn('?leave=3',html);self.assertNotIn('untrusted',html)

    def test_smtp_quit_failure_after_acceptance_does_not_retry(self):
        with patch.dict(os.environ,{'PRESENCE_SMTP_SERVER':'example.com','PRESENCE_SMTP_PORT':'465','PRESENCE_SMTP_PASSWORD':'synthetic'}):
            with patch('backend.leave_notifications.smtplib.SMTP_SSL') as smtp:
                smtp.return_value.__enter__.return_value.send_message.return_value={}
                smtp.return_value.__exit__.side_effect=OSError('synthetic disconnect')
                self.assertEqual(smtp_deliver(Mock()),'sent')
