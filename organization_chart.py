"""Build a safe Graphviz organization chart from API user data."""

from __future__ import annotations

from collections.abc import Mapping


ROLE_COLORS = {
    "Admin": "#8B5CF6",
    "Responsable": "#2563EB",
    "Employe": "#16A34A",
}


def _escape(value: object) -> str:
    return str(value).replace("\\", "\\\\").replace('"', '\\"').replace("\n", " ")


def build_organization_chart(users: list[Mapping[str, object]], group_names: Mapping[int, str]) -> str:
    """Return DOT source for the manager hierarchy, with groups as node metadata."""
    indexed_users = {str(user["username"]): user for user in users if user.get("username")}
    node_ids = {username: f"person_{index}" for index, username in enumerate(sorted(indexed_users))}
    lines = [
        "digraph organization {",
        "rankdir=TB;",
        'graph [pad="0.3", nodesep="0.35", ranksep="0.6", bgcolor="transparent"];',
        'node [shape=box, style="rounded,filled", fontname="sans-serif", fontcolor="white", color="#64748B", margin="0.18,0.12"];',
        'edge [color="#94A3B8", penwidth="1.4"];',
        'root [label="Organisation", fillcolor="#334155", color="#94A3B8"];',
    ]

    for username in sorted(indexed_users):
        user = indexed_users[username]
        role = str(user.get("role") or "Employe")
        group_labels = [group_names.get(int(group_id), str(group_id)) for group_id in user.get("group_ids", [])]
        detail_lines = [role]
        if role == "Employe":
            detail_lines = [f"{role} · {user.get('employee_type') or 'salarie'}"]
        if group_labels:
            detail_lines.append(f"Groupes : {', '.join(group_labels)}")
        label = "\\n".join([_escape(username), *(_escape(line) for line in detail_lines)])
        lines.append(
            f'{node_ids[username]} [label="{label}", fillcolor="{ROLE_COLORS.get(role, "#475569")}"];'
        )

    for username in sorted(indexed_users):
        user = indexed_users[username]
        manager = str(user.get("manager_id") or "")
        if manager and manager in node_ids:
            lines.append(f"{node_ids[manager]} -> {node_ids[username]};")
        else:
            lines.append(f"root -> {node_ids[username]};")

    lines.append("}")
    return "\n".join(lines)
