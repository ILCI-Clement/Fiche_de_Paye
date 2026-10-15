"""ZIP content and failure handling without external network requests."""
from io import BytesIO
import unittest
from unittest.mock import MagicMock, patch
from zipfile import ZipFile
import requests
from document_downloads import build_document_archive


class DocumentDownloadTests(unittest.TestCase):
    def item(self, identifier, filename='test.pdf'):
        return dict(id=identifier, employee_username='employee', document_type='transport', original_filename=filename)

    def response(self, data=b'%PDF-test', status=200):
        response=MagicMock(status_code=status)
        response.__enter__.return_value=response
        response.iter_content.return_value=[data]
        return response

    def test_zip_preserves_selected_content_and_avoids_duplicate_names(self):
        with patch('document_downloads.requests.get', side_effect=[self.response(b'first'),self.response(b'second')]) as get:
            data=build_document_archive([self.item(1),self.item(2)],'https://api.test',{'Authorization':'Bearer test-only'})
        with ZipFile(BytesIO(data)) as archive:
            self.assertEqual(len(set(archive.namelist())),2)
            self.assertEqual([archive.read(name) for name in archive.namelist()],[b'first',b'second'])
        self.assertEqual(get.call_args.kwargs['headers']['Authorization'],'Bearer test-only')

    def test_failed_file_does_not_return_partial_archive(self):
        with patch('document_downloads.requests.get',side_effect=[self.response(),self.response(status=403)]):
            with self.assertRaisesRegex(ValueError,'Aucun ZIP partiel'):
                build_document_archive([self.item(1),self.item(2)],'https://api.test',{})

    def test_size_limit_checks_total_source_bytes(self):
        with patch('document_downloads.MAX_ARCHIVE_BYTES',4),patch('document_downloads.requests.get',return_value=self.response(b'12345')):
            with self.assertRaisesRegex(ValueError,'plusieurs lots'):
                build_document_archive([self.item(1)],'https://api.test',{})

    def test_network_outage_returns_readable_error(self):
        with patch('document_downloads.requests.get',side_effect=requests.Timeout()):
            with self.assertRaisesRegex(ValueError,'indisponible'):
                build_document_archive([self.item(1)],'https://api.test',{})


if __name__=='__main__':unittest.main()
