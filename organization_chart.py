"""Build a safe Graphviz organization chart from API user data."""

from __future__ import annotations

from collections.abc import Mapping
from html import escape


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
        'graph [pad="0.3", nodesep="0.35", ranksep="0.6", bgcolor="transparent", splines="ortho"];',
        'node [shape=box, style="rounded,filled", fontname="sans-serif", fontcolor="white", color="#64748B", margin="0.18,0.12"];',
        'edge [color="#94A3B8", penwidth="1.4"];',
        'root [label="Organisation", fillcolor="#334155", color="#94A3B8"];',
    ]

    for username in sorted(indexed_users):
        user = indexed_users[username]
        tags = [str(tag) for tag in user.get("role_tags", [user.get("role") or "Employe"])]
        role = next((candidate for candidate in ("Admin", "Responsable", "Employe") if candidate in tags), "Employe")
        group_labels = [group_names.get(int(group_id), str(group_id)) for group_id in user.get("group_ids", [])]
        detail_lines = []
        if group_labels:
            detail_lines.append(f"Département : {', '.join(group_labels)}")
        detail_lines.append(username)
        detail_lines.append(" · ".join(tags))
        if "Employe" in tags:
            detail_lines[-1] = f"{' · '.join(tags)} · {user.get('employee_type') or 'salarie'}"
        label = "\\n".join(_escape(line) for line in detail_lines)
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


def build_organization_svg(users: list[Mapping[str, object]], group_names: Mapping[int, str]) -> str:
    """Render a static SVG hierarchy with guaranteed orthogonal connectors."""
    indexed = {str(user["username"]): user for user in users if user.get("username")}
    parents = {
        username: str(user.get("manager_id")) if str(user.get("manager_id") or "") in indexed else None
        for username, user in indexed.items()
    }

    def depth(username: str, seen: set[str] | None = None) -> int:
        seen = seen or set()
        parent = parents[username]
        if not parent or parent in seen:
            return 0
        return depth(parent, seen | {username}) + 1

    levels: dict[int, list[str]] = {}
    for username in sorted(indexed):
        levels.setdefault(depth(username), []).append(username)

    card_width, card_height, gap, top = 230, 120, 38, 48
    max_count = max((len(level) for level in levels.values()), default=1)
    width = max(760, max_count * (card_width + gap) + gap)
    positions: dict[str, tuple[float, float]] = {}
    for level, names in levels.items():
        occupied = len(names) * card_width + (len(names) - 1) * gap
        start_x = (width - occupied) / 2
        for index, username in enumerate(names):
            positions[username] = (start_x + index * (card_width + gap), top + level * 190)
    height = top + (max(levels, default=0) + 1) * 190

    paths: list[str] = []
    for username, parent in parents.items():
        if not parent:
            continue
        parent_x, parent_y = positions[parent]
        child_x, child_y = positions[username]
        parent_center = parent_x + card_width / 2
        child_center = child_x + card_width / 2
        mid_y = parent_y + card_height + (child_y - (parent_y + card_height)) / 2
        paths.append(
            f'<path d="M {parent_center:.0f} {parent_y + card_height:.0f} V {mid_y:.0f} H {child_center:.0f} V {child_y:.0f}" class="connector"/>'
        )

    cards: list[str] = []
    for username in sorted(indexed):
        user = indexed[username]
        x, y = positions[username]
        tags = [str(tag) for tag in user.get("role_tags", [user.get("role") or "Employe"])]
        primary = next((tag for tag in ("Admin", "Responsable", "Employe") if tag in tags), "Employe")
        departments = [group_names.get(int(group_id), str(group_id)) for group_id in user.get("group_ids", [])]
        lines = [f"Département : {', '.join(departments) if departments else 'Non attribué'}", username]
        role_line = " · ".join(tags)
        if "Employe" in tags:
            role_line = f"{role_line} · {user.get('employee_type') or 'salarie'}"
        lines.append(role_line)
        text = "".join(
            f'<text x="{x + card_width / 2:.0f}" y="{y + 34 + index * 27:.0f}" class="line line-{index}">{escape(line)}</text>'
            for index, line in enumerate(lines)
        )
        cards.append(f'<g><rect x="{x:.0f}" y="{y:.0f}" width="{card_width}" height="{card_height}" rx="18" class="card {primary.lower()}"/>{text}</g>')

    return f'''<svg viewBox="0 0 {width} {height}" width="100%" role="img" aria-label="Structure des équipes">
<style>
.connector{{stroke:#94A3B8;stroke-width:3;fill:none}} .card{{stroke:#94A3B8;stroke-width:2}} .admin{{fill:#7C3AED}} .responsable{{fill:#2563EB}} .employe{{fill:#16A34A}}
.line{{fill:#fff;text-anchor:middle;font-family:Arial,sans-serif;font-size:17px}} .line-0{{font-size:13px;fill:#E2E8F0}} .line-1{{font-size:20px;font-weight:700}}
</style>{''.join(paths)}{''.join(cards)}</svg>'''
