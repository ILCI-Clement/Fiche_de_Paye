"""Pure department projection, independent of Streamlit and HTTP."""
import re
import unicodedata

CEO_EMAIL = 'william.tang@univ-ilci.fr'


def person_label(person):
    full_name = str(person.get('full_name') or '').strip()
    if full_name:
        return full_name
    display_name = str(person.get('display_name') or '').strip()
    if display_name:
        return display_name
    email = str(person.get('email') or '').strip()
    return email.partition('@')[0] if re.fullmatch(r'[^@\s]+@[^@\s]+', email) else 'Nom non renseigné'


def normalized_group_name(name):
    name = ''.join(c for c in unicodedata.normalize('NFKD', str(name)) if not unicodedata.combining(c)).casefold().strip()
    return name.removeprefix('groupe ').strip()


def build_organization(users, groups, *, hidden_ids=()):
    hidden = set(hidden_ids)
    people = [p for p in users if p['username'] not in hidden]
    active = [g for g in groups if g.get('is_active')]
    ids = {int(g['id']) for g in active}
    general = next((g for g in active if normalized_group_name(g['name']) == 'direction generale'), None)
    ceo = next((p for p in people if p.get('is_ceo') or str(p.get('email') or '').casefold() == CEO_EMAIL), None)
    unassigned, issues, departments = [], [], []
    for person in people:
        memberships = list(dict.fromkeys(int(i) for i in person.get('group_ids', [])))
        if len(memberships) > 1:
            issues.append(f"{person_label(person)} : plusieurs départements de travail.")
        if not memberships:
            unassigned.append({'Nom': person_label(person), 'Situation': 'Sans département'})
        elif not set(memberships).intersection(ids):
            unassigned.append({'Nom': person_label(person), 'Situation': 'Département inactif'})
    for group in sorted(active, key=lambda g: str(g['name']).casefold()):
        gid = int(group['id'])
        leaders = [p for p in people if gid in p.get('managed_group_ids', [])]
        members = [p for p in people if gid in p.get('group_ids', [])]
        # A leader can lead several departments and is listed within each card.
        combined = {p['username']: p for p in members + leaders}
        if general and gid == int(general['id']) and ceo:
            combined[ceo['username']] = ceo
        elif ceo:
            combined.pop(ceo['username'], None)
        if len(leaders) > 1:
            issues.append(f"{group['name']} : plusieurs responsables, affectation à corriger.")
        departments.append({
            'id': gid, 'name': str(group['name']),
            'leaders': [person_label(p) for p in leaders],
            'members': [person_label(p) for p in sorted(combined.values(), key=lambda p: person_label(p).casefold())],
        })
    if not general:
        issues.append('Direction Générale est absente des Groupes actifs; le nœud central est un repère de présentation.')
    if not ceo:
        issues.append('Le compte CEO configuré est absent du périmètre retourné par le serveur.')
    return {'ceo': person_label(ceo) if ceo else 'Nom non renseigné',
            'general_id': int(general['id']) if general else None,
            'departments': departments, 'unassigned': unassigned, 'issues': issues,
            'people_count': len(people)}

