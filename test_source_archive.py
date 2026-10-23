"""Frozen UI integrity and retained calendar compatibility after source separation."""
import ast
import hashlib
import json
from pathlib import Path
import unittest
import zipfile


class SourceArchiveTests(unittest.TestCase):
    def test_archive_and_removed_sources_match_recorded_hashes(self):
        root = Path(__file__).parent
        folder = root / 'archive/streamlit'
        manifest = json.loads((folder / 'manifest.json').read_text(encoding='utf-8'))
        content = (folder / manifest['archive']).read_bytes()
        self.assertEqual(len(content), manifest['bytes'])
        self.assertEqual(hashlib.sha256(content).hexdigest(), manifest['sha256'])
        with zipfile.ZipFile(folder / manifest['archive']) as archive:
            self.assertIsNone(archive.testzip())
            for name, digest in manifest['removed_files'].items():
                self.assertEqual(hashlib.sha256(archive.read(name)).hexdigest(), digest, name)
                self.assertFalse((root / name).exists(), name)
            for name in archive.namelist():
                self.assertNotIn('.local-test/', name)
                self.assertNotIn('.venv/', name)
                self.assertNotIn('.git/', name)
                self.assertFalse(name.endswith('secrets.toml'), name)

    def test_retained_calendar_functions_match_frozen_implementation(self):
        root = Path(__file__).parent
        with zipfile.ZipFile(root / 'archive/streamlit/source.zip') as archive:
            original = ast.parse(archive.read('calendar_view.py').decode('utf-8-sig'))
        retained = ast.parse((root / 'calendar_view.py').read_text(encoding='utf-8-sig'))
        functions = {item.name: ast.dump(item) for item in original.body if isinstance(item, ast.FunctionDef)}
        for item in retained.body:
            if isinstance(item, ast.FunctionDef):
                self.assertEqual(ast.dump(item), functions[item.name], item.name)
