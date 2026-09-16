import unittest

from access_control import can_edit_employee, can_manage_employee, normalize_role


class AccessControlTests(unittest.TestCase):
    def test_legacy_admin_flag_keeps_admin_role(self):
        self.assertEqual(normalize_role({"is_admin": True}), "Admin")

    def test_admin_manages_any_employee(self):
        self.assertTrue(can_manage_employee({"role": "Admin"}, {"group_ids": ["other"]}))

    def test_responsable_manages_employee_in_any_owned_group(self):
        user = {"role": "Responsable", "id": "manager-1", "managed_group_ids": ["RH", "Finance"]}
        employee = {"manager_id": "manager-2", "group_ids": ["Finance", "R&D"]}
        self.assertTrue(can_manage_employee(user, employee))

    def test_responsable_manages_direct_report_without_matching_group(self):
        user = {"role": "Responsable", "id": "manager-1", "managed_group_ids": []}
        employee = {"manager_id": "manager-1", "group_ids": ["other"]}
        self.assertTrue(can_manage_employee(user, employee))

    def test_responsable_cannot_manage_out_of_scope_employee(self):
        user = {"role": "Responsable", "id": "manager-1", "managed_group_ids": ["RH"]}
        employee = {"manager_id": "manager-2", "group_ids": ["Finance"]}
        self.assertFalse(can_manage_employee(user, employee))

    def test_employe_is_read_only(self):
        user = {"role": "Employe", "id": "employee-1"}
        employee = {"manager_id": "employee-1", "group_ids": ["RH"]}
        self.assertFalse(can_edit_employee(user, employee))


if __name__ == "__main__":
    unittest.main()
