"""Tests for the Presence API storage boundary."""

from datetime import date
from unittest import TestCase
from unittest.mock import Mock

import requests

from presence_storage import PresenceApiStorage, PresenceStorageError


class PresenceApiStorageTests(TestCase):
    def create_storage(self, client):
        return PresenceApiStorage(
            "https://presence.example.test/api/",
            "test-token",
            http_client=client,
        )

    def test_loads_raw_workspace_and_encodes_username(self):
        response = Mock()
        response.json.return_value = {"legacy_key": "preserved"}
        client = Mock()
        client.get.return_value = response

        workspace = self.create_storage(client).load_workspace("Jean Dupont/a")

        self.assertEqual(workspace["schema_version"], 2)
        self.assertEqual(workspace["legacy_key"], "preserved")
        client.get.assert_called_once_with(
            "https://presence.example.test/api/get-config/Jean%20Dupont%2Fa",
            headers={
                "Authorization": "Bearer test-token",
                "Content-Type": "application/json",
            },
            timeout=10,
        )

    def test_rejects_non_object_response(self):
        response = Mock()
        response.json.return_value = []
        client = Mock()
        client.get.return_value = response

        with self.assertRaises(PresenceStorageError):
            self.create_storage(client).load_workspace("tester")

    def test_network_failure_is_wrapped(self):
        client = Mock()
        client.get.side_effect = requests.ConnectionError("offline")

        with self.assertRaisesRegex(PresenceStorageError, "API Presence"):
            self.create_storage(client).load_workspace("tester")

    def test_save_serializes_dates_without_mutating_workspace(self):
        response = Mock()
        client = Mock()
        client.post.return_value = response
        workspace = {
            "schema_version": 2,
            "people": [{"start_date": date(2026, 9, 1)}],
            "forms": [],
            "trash": [],
        }

        self.create_storage(client).save_workspace("tester", workspace)

        payload = client.post.call_args.kwargs["json"]
        self.assertEqual(payload["people"][0]["start_date"], "2026-09-01")
        self.assertEqual(workspace["people"][0]["start_date"], date(2026, 9, 1))
        response.raise_for_status.assert_called_once_with()

    def test_configuration_is_required(self):
        with self.assertRaises(ValueError):
            PresenceApiStorage("", "test-token")
        with self.assertRaises(ValueError):
            PresenceApiStorage("https://presence.example.test", "")
