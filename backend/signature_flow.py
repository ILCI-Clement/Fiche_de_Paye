"""Resolve both attendance signers from authorized personnel accounts."""
from fastapi import HTTPException
import re


def signer_display_name(name, email):
    """Keep mailbox addresses out of the provider's signer-name field."""
    value = str(name or '').strip() or str(email or '').strip()
    value = re.sub(r'[^\s<>(),;]+@[^\s<>(),;]+', lambda match: match.group(0).split('@', 1)[0], value)
    value = ' '.join(value.strip(' <>').split())
    if not value or '@' in value:
        raise HTTPException(400, 'Renseignez un nom de signataire valide, distinct de son e-mail.')
    return value


def resolve_manager(recipient_email, actor, connect, load_user, access):
    db = connect()
    try:
        with db.cursor() as cursor:
            cursor.execute('SELECT username FROM users WHERE LOWER(email)=LOWER(%s)', (str(recipient_email).strip(),))
            matches = cursor.fetchall()
            if len(matches) != 1:
                raise HTTPException(400, 'Le salarié doit être associé à un compte unique enregistré.')
            employee = access(cursor, actor, matches[0]['username'], edit=False)
            manager_username = employee.get('manager_username')
            if not manager_username or manager_username == employee['username']:
                raise HTTPException(400, 'Configurez le responsable direct de cette personne dans Administration avant l’envoi.')
            manager = load_user(cursor, manager_username)
            if not manager or not ({'Admin','Responsable'} & set(manager['role_tags'])) or not str(manager.get('email') or '').strip():
                raise HTTPException(400, 'Le responsable direct doit avoir un compte valide et une adresse e-mail.')
            if manager['email'].strip().casefold() == str(recipient_email).strip().casefold():
                raise HTTPException(400, 'Le salarié et le responsable doivent avoir des adresses e-mail distinctes.')
            name = ' '.join(str(manager.get(key) or '').strip() for key in ('first_name','last_name')).strip()
            return {'name': signer_display_name(name or manager['username'], manager['email']), 'email':manager['email'].strip()}
    finally:
        db.close()


def signing_people(employee_name, employee_email, manager_name, manager_email):
    return [
        {'name':signer_display_name(employee_name, employee_email),'email':employee_email,'order':1,'role':'SIGNER'},
        {'name':signer_display_name(manager_name, manager_email),'email':manager_email,'order':2,'role':'SIGNER'},
    ]


def attendance_signature_fields(*, intern=False):
    """Exact template anchors; order-2 SIGNATURE enables the manager's handwriting."""
    return [
        {'name': 'sig_salarie', 'type': 'SIGNATURE', 'party': 0,
         'anchor': 'Signature du stagiaire :' if intern else 'Signature du salarié :',
         'anchor_scope': 'document', 'dx': 0, 'dy': 16, 'width': 150, 'height': 70},
        {'name': 'sig_manager', 'type': 'SIGNATURE', 'party': 1,
         'anchor': 'Signature et Cachet Entreprise Cliente :',
         'anchor_scope': 'document', 'dx': 0, 'dy': 16, 'width': 150, 'height': 70},
    ]
