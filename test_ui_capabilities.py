import unittest
from fastapi import FastAPI
from fastapi.testclient import TestClient
from backend.ui_capabilities import register_ui_capabilities


class UiCapabilitiesTests(unittest.TestCase):
    def client(self, *, local=False, mail=True):
        app = FastAPI()
        app.state.local_snapshot = local
        register_ui_capabilities(app, mail_configured=lambda: mail)
        client = TestClient(app)
        self.addCleanup(client.close)
        return client

    def test_local_snapshot_disables_external_actions(self):
        self.assertEqual(self.client(local=True).get('/ui-capabilities').json(), {
            'environment': 'local-snapshot', 'account_creation': False, 'mail_enabled': False,
        })

    def test_server_exposes_only_boolean_configuration(self):
        response = self.client().get('/ui-capabilities')
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.json(), {'environment': 'server', 'account_creation': True, 'mail_enabled': True})

    def test_missing_smtp_configuration_disables_mail(self):
        self.assertFalse(self.client(mail=False).get('/ui-capabilities').json()['mail_enabled'])

    def test_capabilities_cannot_be_changed_by_query(self):
        response = self.client(local=True).get('/ui-capabilities?account_creation=true&mail_enabled=true')
        self.assertFalse(response.json()['account_creation'])
        self.assertFalse(response.json()['mail_enabled'])


if __name__ == '__main__':
    unittest.main()
