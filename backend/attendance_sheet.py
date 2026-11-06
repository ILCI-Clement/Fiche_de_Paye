"""One constructor for month/contract-bound sheets, including leave approvals."""
import copy

from contract_periods import bind_contract
from fiche_serialization import serialize_dates
from planning import default_planning

from .person_names import person_display_name


def new_sheet(account, manager, contract, month, year, records):
    name = person_display_name(account)
    record = {'nom': name, 'nom_stagiaire': name, 'prenom_stagiaire': '',
              'email_employe': account['email'], 'account_username': account['username'],
              'responsable': person_display_name(manager) if manager else '',
              'email_responsable': (manager or {}).get('email', ''),
              'nb_jours': 0, 'nb_heures_jour': 0.0, 'taux_horaire': 0.0,
              'facture_mensuelle': 0.0, 'transport': '', 'taux': 0.0,
              'vacances': [], 'absences': [], 'arret': [], 'calendar_overrides': {},
              'planning_detail': default_planning()}
    prior = next((e for e in reversed(records) if e.get('account_username') == account['username']), None)
    if prior:
        for field in ('nom', 'nom_stagiaire', 'prenom_stagiaire', 'responsable', 'email_responsable', 'planning_detail'):
            if prior.get(field):
                record[field] = copy.deepcopy(prior[field])
        if prior.get('type') == ('Stagiaire' if contract['employee_type'] == 'stagiaire' else 'Salarié'):
            for field in ('taux_horaire', 'nb_heures_jour', 'transport', 'taux'):
                if field in prior:
                    record[field] = copy.deepcopy(prior[field])
    return serialize_dates(bind_contract(record, contract, month, year))
