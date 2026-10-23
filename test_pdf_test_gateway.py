import json
import tempfile
import time
import unittest
from unittest.mock import Mock, patch
from pathlib import Path
from fastapi.testclient import TestClient
from backend.pdf_test_gateway import create_pdf_gateway


class PdfGatewayTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.token = 'a' * 43
        (self.root / f'{self.token}.pdf').write_bytes(b'%PDF-unit-fixture')
        self.metadata = {'pdf_name': f'{self.token}.pdf', 'expires_at': time.time() + 60}
        self.write_metadata()
        self.client = TestClient(create_pdf_gateway(self.root))
        self.addCleanup(self.client.close)

    def write_metadata(self):
        (self.root / f'{self.token}.json').write_text(json.dumps(self.metadata))

    def test_only_valid_pdf_is_readable(self):
        response = self.client.get(f'/esign/documents/{self.token}')
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.content, b'%PDF-unit-fixture')
        self.assertIn('no-store', response.headers['cache-control'])

    def test_private_routes_are_absent(self):
        for path in ['/login', '/list-users', '/transport-receipts', '/attendance/workspace', '/docs', '/openapi.json', '/', '/esign/documents/']:
            self.assertEqual(self.client.get(path).status_code, 404)

    def test_write_methods_are_blocked(self):
        for method in ['POST', 'PUT', 'PATCH', 'DELETE']:
            self.assertEqual(self.client.request(method, f'/esign/documents/{self.token}').status_code, 405)

    def test_expired_or_invalid_metadata_denies_access(self):
        for expiry in [time.time() - 1, float('inf'), float('nan')]:
            self.metadata['expires_at'] = expiry
            self.write_metadata()
            self.assertEqual(self.client.get(f'/esign/documents/{self.token}').status_code, 404)

    def test_metadata_cannot_select_different_file(self):
        self.metadata['pdf_name'] = '../private.pdf'
        self.write_metadata()
        self.assertEqual(self.client.get(f'/esign/documents/{self.token}').status_code, 404)

    def test_token_and_query_are_restricted(self):
        for token in ['a', 'a' * 42, self.token + '.json', self.token + '?download=1']:
            self.assertEqual(self.client.get(f'/esign/documents/{token}').status_code, 404)

    def test_head_has_no_body(self):
        response = self.client.head(f'/esign/documents/{self.token}')
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.content, b'')

    def test_callback_only_forwards_to_fixed_local_api_without_client_headers(self):
        upstream = Mock(); upstream.status = 204
        upstream.__enter__ = Mock(return_value=upstream); upstream.__exit__ = Mock(return_value=False)
        with patch('backend.pdf_test_gateway.urllib.request.build_opener') as opener:
            opener.return_value.open.return_value = upstream
            response = self.client.post('/esign/callback', json={'reference_id': 'synthetic'},
                                        headers={'Authorization': 'Bearer synthetic-do-not-forward'})
            self.assertEqual(response.status_code, 204)
            request = opener.return_value.open.call_args.args[0]
            self.assertEqual(request.full_url, 'http://127.0.0.1:8013/esign/callback')
            self.assertIsNone(request.get_header('Authorization'))
            self.assertEqual(request.get_method(), 'POST')

    def test_oversized_and_query_callbacks_never_reach_api(self):
        with patch('backend.pdf_test_gateway.urllib.request.build_opener') as opener:
            self.assertEqual(self.client.post('/esign/callback',content=b'x'*32769).status_code,413)
            self.assertEqual(self.client.post('/esign/callback?target=http://example.test',json={}).status_code,404)
            opener.assert_not_called()


if __name__ == '__main__':
    unittest.main()
