"""Shared role and organization access rules for the Streamlit client."""

from __future__ import annotations

from collections.abc import Mapping


VALID_ROLES = {"Admin", "Responsable", "Employe"}
ROLE_PRIORITY = ("Admin", "Responsable", "Employe")


def role_tags(user: Mapping[str, object]) -> set[str]:
    tags = user.get("role_tags")
    if isinstance(tags, (list, tuple, set)):
        recognized = {str(tag) for tag in tags if str(tag) in VALID_ROLES}
        if recognized:
            return recognized
    role = user.get("role")
    if role in VALID_ROLES:
        return {str(role)}
    return {"Admin"} if user.get("is_admin") else {"Responsable"}


def normalize_role(user: Mapping[str, object]) -> str | None:
    """Return the existing role name, with compatibility for legacy is_admin data."""
    return next((role for role in ROLE_PRIORITY if role in role_tags(user)), None)


def managed_group_ids(user: Mapping[str, object]) -> set[str]:
    """Read group identifiers assigned to a Responsable from either supported shape."""
    groups = user.get("managed_group_ids", user.get("groups", []))
    if isinstance(groups, str):
        return {groups}
    if isinstance(groups, (list, tuple, set)):
        return {str(group) for group in groups}
    return set()


def employee_group_ids(employee: Mapping[str, object]) -> set[str]:
    """Read an employee's groups while accepting the singular legacy field."""
    groups = employee.get("group_ids", employee.get("groups", []))
    if isinstance(groups, str):
        return {groups}
    if isinstance(groups, (list, tuple, set)):
        return {str(group) for group in groups}
    return set()


def can_manage_employee(user: Mapping[str, object], employee: Mapping[str, object]) -> bool:
    """Apply the documented client-side visibility rule.

    The Presence API must enforce the same rule server-side; this helper only keeps
    Streamlit navigation and filtering consistent with the API contract.
    """
    role = normalize_role(user)
    if role == "Admin":
        return True
    if role != "Responsable":
        return False

    employee_manager = employee.get("manager_id", employee.get("responsable_id"))
    if employee_manager is not None and str(employee_manager) == str(user.get("id", user.get("username", user.get("name", "")))):
        return True
    return bool(managed_group_ids(user) & employee_group_ids(employee))


def can_edit_employee(user: Mapping[str, object], employee: Mapping[str, object]) -> bool:
    """Editing uses the same scope as management; Employe remains read-only."""
    return can_manage_employee(user, employee)
