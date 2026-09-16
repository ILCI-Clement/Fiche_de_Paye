import unittest

from organization_chart import build_organization_chart, build_organization_png, build_organization_svg


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

    def test_svg_uses_only_orthogonal_connector_segments(self):
        svg = build_organization_svg(
            [{"username": "manager", "role": "Responsable"}, {"username": "employee", "role": "Employe", "manager_id": "manager", "group_ids": [2]}],
            {2: "R&D"},
        )
        self.assertIn(' V ', svg)
        self.assertIn(' H ', svg)
        self.assertIn('Département : R&amp;D', svg)

    def test_png_chart_is_generated(self):
        image = build_organization_png([{"username": "admin", "role": "Admin"}], {})
        self.assertTrue(image.startswith(b"\x89PNG\r\n\x1a\n"))


if __name__ == "__main__":
    unittest.main()
