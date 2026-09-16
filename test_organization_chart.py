import unittest

from organization_chart import build_organization_chart


class OrganizationChartTests(unittest.TestCase):
    def test_renders_direct_manager_and_group_metadata(self):
        dot = build_organization_chart(
            [
                {"username": "admin", "role": "Admin", "group_ids": []},
                {"username": "alice", "role": "Employe", "employee_type": "salarie", "manager_id": "admin", "group_ids": [3]},
            ],
            {3: "Paris"},
        )
        self.assertIn('root -> person_0;', dot)
        self.assertIn('person_0 -> person_1;', dot)
        self.assertIn('Département : Paris', dot)
        self.assertIn('splines="ortho"', dot)

    def test_escapes_user_provided_text(self):
        dot = build_organization_chart([{"username": 'A " user', "role": "Admin"}], {})
        self.assertIn('A \\" user', dot)


if __name__ == "__main__":
    unittest.main()
