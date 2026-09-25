"""Focused regression tests for backend data-integrity helpers."""

from __future__ import annotations

import json
import os
import unittest


os.environ.setdefault("PRESENCE_API_TOKEN", "test-api-token")
os.environ.setdefault("PRESENCE_DB_HOST", "localhost")
os.environ.setdefault("PRESENCE_DB_USER", "test")
os.environ.setdefault("PRESENCE_DB_PASSWORD", "test")
os.environ.setdefault("PRESENCE_DB_DATABASE", "test")
os.environ.setdefault("PRESENCE_SMTP_SERVER", "localhost")
os.environ.setdefault("PRESENCE_SMTP_PORT", "465")
os.environ.setdefault("PRESENCE_SMTP_USER", "test@example.com")
os.environ.setdefault("PRESENCE_SMTP_PASSWORD", "test")
os.environ.setdefault("PRESENCE_APP_URL", "https://example.test")

from backend import main


class SavedFicheCursor:
    def __init__(self, rows: list[dict]) -> None:
        self.rows = rows
        self.commands: list[tuple[str, tuple | None]] = []

    def execute(self, query: str, params: tuple | None = None) -> None:
        self.commands.append((query, params))

    def fetchall(self) -> list[dict]:
        return self.rows


class ClaimCursor:
    def __init__(self, rowcount: int) -> None:
        self.rowcount = rowcount
        self.params: tuple | None = None

    def execute(self, _: str, params: tuple | None = None) -> None:
        self.params = params


class BackendHelperTests(unittest.TestCase):
    def test_fiche_config_rejects_more_than_one_hundred_employees(self) -> None:
        config = {"employes_data": [{} for _ in range(main.MAX_FICHE_EMPLOYEES + 1)]}
        with self.assertRaises(main.HTTPException) as error:
            main.serialize_fiche_config(config)
        self.assertEqual(error.exception.status_code, 400)

    def test_username_migration_updates_saved_fiche_links_and_emails(self) -> None:
        cursor = SavedFicheCursor(
            [
                {
                    "user_id": "manager",
                    "form_content": json.dumps(
                        {
                            "employes_data": [
                                {
                                    "account_username": "old-name",
                                    "responsable": "old-name",
                                    "email_employe": "old@example.test",
                                    "email_responsable": "old@example.test",
                                }
                            ]
                        }
                    ),
                }
            ]
        )

        main.migrate_username_in_saved_fiches(
            cursor,
            "old-name",
            "new-name",
            "old@example.test",
            "new@example.test",
        )

        updates = [command for command in cursor.commands if command[0].startswith("UPDATE Presence")]
        self.assertEqual(len(updates), 1)
        migrated = json.loads(updates[0][1][0])
        employee = migrated["employes_data"][0]
        self.assertEqual(employee["account_username"], "new-name")
        self.assertEqual(employee["responsable"], "new-name")
        self.assertEqual(employee["email_employe"], "new@example.test")
        self.assertEqual(employee["email_responsable"], "new@example.test")

    def test_reminder_claim_uses_a_single_insert_before_delivery(self) -> None:
        cursor = ClaimCursor(rowcount=1)
        contract = {"reminder_key": "employee", "contract_end_date": "2026-10-01"}

        self.assertTrue(main.claim_contract_end_reminder(cursor, contract, "admin@example.test"))
        self.assertEqual(cursor.params, ("employee", "2026-10-01", "admin@example.test"))

    def test_existing_reminder_claim_is_not_sent_again(self) -> None:
        cursor = ClaimCursor(rowcount=0)
        contract = {"reminder_key": "employee", "contract_end_date": "2026-10-01"}

        self.assertFalse(main.claim_contract_end_reminder(cursor, contract, "admin@example.test"))


if __name__ == "__main__":
    unittest.main()
