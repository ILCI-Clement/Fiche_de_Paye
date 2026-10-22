"""Retained API must import in its deployed flat-module layout."""
import importlib.util
from pathlib import Path
import sys
import unittest
from unittest.mock import patch


class FlatBackendImportTests(unittest.TestCase):
    def test_legacy_shared_modules_support_flat_imports(self):
        root = Path(__file__).parent / 'backend'
        for name in ('team_directory', 'attendance_permissions', 'account_purge'):
            with self.subTest(module=name), patch.object(sys, 'path', [str(root), *sys.path]):
                spec = importlib.util.spec_from_file_location('flat_' + name, root / (name + '.py'))
                module = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(module)
                self.assertEqual(module.__package__, '')


if __name__ == '__main__':
    unittest.main()
