import calendar
from copy import deepcopy
from datetime import date, datetime

import streamlit as st

from app_core import (
    WEEK_DAYS,
    create_form,
    create_person,
    default_schedule,
    ensure_workspace,
    find_duplicate_form,
    move_to_trash,
    person_display_name,
    restore_from_trash,
    validate_form,
    import_legacy_forms,
)
from export_service import (
    EXPORT_LABELS,
    prepare_exports,
    export_fingerprint,
)
from attendance import monthly_attendance, format_minutes
from presence_storage import PresenceApiStorage, PresenceStorageError, serialize_dates


EXCEPTION_LABELS = {
    "work": "Travail normal",
    "rest": "Repos",
    "paid_leave": "Congé payé (CP)",
    "absence": "Absence (ABS)",
    "sick_leave": "Arrêt maladie (AM)",
    "public_holiday": "Jour férié / fermeture",
    "other": "Autre",
}
FORM_STATE_LABELS = {
    "draft": "Brouillon",
    "complete": "Terminé",
    "exported": "Exporté",
    "archived": "Archivé",
}


def create_server_storage(api_url, headers):
    authorization = headers.get("Authorization", "")
    prefix = "Bearer "
    token = authorization[len(prefix):] if authorization.startswith(prefix) else ""
    return PresenceApiStorage(api_url, token)


def load_workspace(api_url, headers, username):
    cache_key = f"workspace_{username}"
    if cache_key in st.session_state:
        return st.session_state[cache_key]
    try:
        workspace = create_server_storage(api_url, headers).load_workspace(username)
    except (PresenceStorageError, ValueError) as error:
        st.error(str(error))
        st.stop()
    st.session_state[cache_key] = workspace
    return workspace


def save_workspace(api_url, headers, username, workspace):
    if st.session_state.get('_local_preview'):
        from preview_storage import save_preview_workspace
        try:
            save_preview_workspace(serialize_dates(workspace))
            st.success('Données de test sauvegardées sur cet ordinateur. Aucun envoi au serveur.')
        except Exception as error:
            st.error(f'Sauvegarde locale impossible : {error}')
        return
    errors = [f"{get_form_name(form)} : {error}" for form in workspace['forms'] for error in validate_form(form)]
    if errors:
        for error in errors:
            st.error(error)
        return
    try:
        create_server_storage(api_url, headers).save_workspace(username, workspace)
        st.success("Les données ont été sauvegardées sur le serveur.")
    except (PresenceStorageError, ValueError) as error:
        st.error(str(error))


def update_home_metrics(workspace):
    forms = workspace["forms"]
    st.session_state["draft_count"] = sum(item["state"] == "draft" for item in forms)
    st.session_state["exported_count"] = sum(
        item["state"] == "exported" and not item.get("archived") for item in forms
    )
    st.session_state["archived_count"] = sum(item.get("archived") for item in forms)


def get_form_name(form):
    return " ".join(
        part for part in (form.get("last_name", "").upper(), form.get("first_name", "")) if part
    ) or "Sans nom"


def mark_downloaded(workspace, package, formats):
    records = {form['id']: form for form in workspace['forms']}
    if not all(item in records for item in package['ids']):
        return
    forms = [records[item] for item in package['ids']]
    if export_fingerprint(forms) != package['fingerprint']:
        return
    for form in forms:
        if not form.get('archived'):
            form['state'] = 'exported'
        form['exported_formats'] = sorted(set(form.get('exported_formats', [])) | set(formats))


def filter_forms(forms, key):
    """Apply filters without changing the records stored in the workspace."""
    columns = st.columns(3)
    query = columns[0].text_input('Rechercher par nom', key=f'{key}_query').casefold().strip()
    person_type = columns[1].selectbox('Type', ['all', 'employee', 'intern'],
        format_func=lambda value: {'all': 'Tous', 'employee': 'Salarié', 'intern': 'Stagiaire'}[value], key=f'{key}_type')
    periods = sorted({f"{form['year']}-{form['month']:02d}" for form in forms}, reverse=True)
    period = columns[2].selectbox('Période', ['Toutes'] + periods, key=f'{key}_period')
    return [form for form in forms if (not query or query in get_form_name(form).casefold())
        and (person_type == 'all' or person_type == form['person_type'])
        and (period == 'Toutes' or period == f"{form['year']}-{form['month']:02d}")]


@st.dialog('Supprimer définitivement cette fiche ?')
def confirm_permanent_delete(username, form_id):
    st.write('Cette opération est irréversible. La fiche sera retirée de la corbeille ; les fichiers déjà téléchargés ne sont pas concernés.')
    if st.button('Confirmer la suppression définitive', type='primary'):
        workspace = st.session_state[f'workspace_{username}']
        workspace['trash'][:] = [form for form in workspace['trash'] if form['id'] != form_id]
        st.session_state.pop(f'pending_delete_{username}', None)
        st.rerun()
    if st.button('Annuler'):
        st.session_state.pop(f'pending_delete_{username}', None)
        st.rerun()


def render_schedule_editor(person):
    schedule = person.setdefault("default_schedule", default_schedule())
    with st.container():
        st.markdown('**Planning hebdomadaire par défaut**')
        st.caption("Les modifications s’appliquent uniquement aux fiches créées ultérieurement.")
        for day in WEEK_DAYS:
            day_schedule = schedule.setdefault(day, default_schedule()[day])
            columns = st.columns([1.4, 1.2, 1.4, 1.4, 1.4])
            day_schedule["active"] = columns[0].checkbox(
                day, value=day_schedule["active"], key=f"schedule_active_{person['id']}_{day}"
            )
            if day_schedule["active"]:
                day_schedule["morning_start"] = columns[1].text_input(
                    "Matin", day_schedule["morning_start"], key=f"schedule_morning_start_{person['id']}_{day}"
                )
                day_schedule["morning_end"] = columns[2].text_input(
                    "Fin matin", day_schedule["morning_end"], key=f"schedule_morning_end_{person['id']}_{day}"
                )
                day_schedule["afternoon_start"] = columns[3].text_input(
                    "Après-midi", day_schedule["afternoon_start"], key=f"schedule_afternoon_start_{person['id']}_{day}"
                )
                day_schedule["afternoon_end"] = columns[4].text_input(
                    "Fin après-midi", day_schedule["afternoon_end"], key=f"schedule_afternoon_end_{person['id']}_{day}"
                )


def render_details_editor(record, key, is_form=False):
    """Edit a person or an independent form snapshot without reverse synchronization."""
    before = deepcopy(record)
    fields = record.setdefault('person_snapshot', {}) if is_form else record
    type_key = 'person_type' if is_form else 'type'
    columns = st.columns(2)
    if not is_form:
        record[type_key] = columns[0].selectbox('Type de personne', ['employee', 'intern'],
            index=['employee', 'intern'].index(record[type_key]), format_func=lambda value: 'Salarié' if value == 'employee' else 'Stagiaire', key=f'{key}_type')
    record['last_name'] = columns[0].text_input('Nom', record.get('last_name', ''), key=f'{key}_last')
    record['first_name'] = columns[1].text_input('Prénom', record.get('first_name', ''), key=f'{key}_first')
    fields['supervisor_name'] = columns[0].text_input('Responsable', fields.get('supervisor_name', ''), key=f'{key}_supervisor')
    fields['supervisor_email'] = columns[1].text_input('E-mail du responsable', fields.get('supervisor_email', ''), key=f'{key}_email')
    if record[type_key] == 'employee':
        fields['permanent_contract'] = columns[1].checkbox('Contrat CDI', value=bool(fields.get('permanent_contract')), key=f'{key}_permanent')
    for field, column, label in [('start_date', columns[0], 'Début du contrat / stage'), ('end_date', columns[1], 'Fin du contrat / stage')]:
        if field == 'end_date' and fields.get('permanent_contract') and record[type_key] == 'employee':
            continue
        raw = fields.get(field)
        try:
            parsed = date.fromisoformat(str(raw)) if raw else None
        except ValueError:
            parsed = None
        result = column.date_input(label, value=parsed, format='DD/MM/YYYY', key=f'{key}_{field}')
        fields[field] = result.isoformat() if result else None
    if record[type_key] == 'intern':
        st.markdown('**Indemnité de stage**')
        st.caption('Base de calcul saisie manuellement, comme dans la version précédente. Le calendrier ne modifie pas automatiquement les jours indemnisés.')
        for index, (field, label, maximum) in enumerate([
            ('hourly_rate', 'Taux horaire (€)', None), ('day_count', 'Jours indemnisés', 31.0),
            ('daily_hours', 'Heures indemnisées / jour', 24.0), ('transport_cost', 'Facture transport (€)', None),
            ('transport_rate', 'Remboursement transport (%)', 100.0)]):
            fields[field] = columns[index % 2].number_input(label, min_value=0.0, max_value=maximum,
                value=float(fields.get(field, 0.0)), step=0.5, key=f'{key}_{field}')
        fields['transport'] = st.text_input('Transport', fields.get('transport', ''), key=f'{key}_transport')
    if is_form and before != record and not record.get('archived'):
        record['state'] = 'draft'


def render_calendar_editor(form):
    before = deepcopy(form.get('exceptions', {}))
    st.subheader("Calendrier mensuel")
    st.caption('Cliquez sur une date, puis modifiez directement le matin ou l’après-midi dans le panneau placé au-dessus du calendrier.')
    exception_options = list(EXCEPTION_LABELS)
    try:
        attendance = monthly_attendance(form)
    except ValueError as error:
        st.error(str(error))
        return
    st.metric('Heures du mois', format_minutes(sum(day['minutes'] for day in attendance)))

    selected_day_key = f"selected_day_{form['id']}"
    today = date.today()
    default_day = today.day if (today.year, today.month) == (form['year'], form['month']) else 1
    selected_day = int(st.session_state.get(selected_day_key, default_day))
    if not 1 <= selected_day <= len(attendance):
        selected_day = default_day
        st.session_state[selected_day_key] = selected_day

    current_date = date(form['year'], form['month'], selected_day).isoformat()
    current_attendance = attendance[selected_day - 1]
    base_form = deepcopy(form)
    base_day = base_form.setdefault('exceptions', {}).get(current_date, {})
    for half_day in ('morning', 'afternoon'):
        base_day.pop(half_day, None)
    if not base_day:
        base_form['exceptions'].pop(current_date, None)
    base_attendance = monthly_attendance(base_form)[selected_day - 1]
    updated_exceptions = deepcopy(before)
    updated_day = updated_exceptions.setdefault(current_date, {})

    with st.container(border=True):
        st.markdown(f"#### Modifier le {selected_day:02d}/{form['month']:02d}/{form['year']} — {current_attendance['weekday']}")
        edit_columns = st.columns(2)
        for half_day, column, label in (
            ('morning', edit_columns[0], 'Matin'),
            ('afternoon', edit_columns[1], 'Après-midi'),
        ):
            existing_item = deepcopy(before.get(current_date, {}).get(half_day, {}))
            current_type = current_attendance[half_day]['type']
            selected_type = column.radio(
                f'Statut — {label}',
                exception_options,
                index=exception_options.index(current_type),
                format_func=lambda value: EXCEPTION_LABELS[value],
                key=f"calendar_{form['id']}_{current_date}_{half_day}",
                horizontal=True,
            ) or current_type
            updated_item = {'type': selected_type}
            if selected_type == 'paid_leave':
                exam_leave = column.checkbox(
                    'Examen alternance',
                    value=bool(existing_item.get('exam_leave')),
                    key=f"exam_{form['id']}_{current_date}_{half_day}",
                )
                if exam_leave:
                    updated_item['exam_leave'] = True
            if selected_type == 'other':
                updated_item['label'] = column.text_input(
                    'Précision / raison',
                    existing_item.get('label', ''),
                    key=f"other_label_{form['id']}_{current_date}_{half_day}",
                )
                updated_item['hours'] = column.number_input(
                    'Heures',
                    min_value=0.0,
                    max_value=24.0,
                    value=float(existing_item.get('hours', 0.0)),
                    step=0.5,
                    key=f"other_hours_{form['id']}_{current_date}_{half_day}",
                )

            if updated_item == {'type': base_attendance[half_day]['type']}:
                updated_day.pop(half_day, None)
            else:
                updated_day[half_day] = updated_item
        if not updated_day:
            updated_exceptions.pop(current_date, None)

    weekday_headers = st.columns(7)
    for index, weekday in enumerate(WEEK_DAYS):
        weekday_headers[index].markdown(f"**{weekday[:3]}**")
    for week in calendar.monthcalendar(form['year'], form['month']):
        columns = st.columns(7)
        for index, number in enumerate(week):
            if number:
                day = attendance[number-1]
                columns[index].button(
                    f"{number:02d}\n\nM : {day['morning']['label']}\n\nA : {day['afternoon']['label']}",
                    key=f"select_day_{form['id']}_{number}",
                    type='primary' if number == selected_day else 'secondary',
                    help='Cliquer pour modifier cette journée',
                    use_container_width=True,
                    on_click=st.session_state.__setitem__,
                    args=(selected_day_key, number),
                )

    if before != updated_exceptions:
        form['exceptions'] = updated_exceptions
        if not form.get('archived'):
            form['state'] = 'draft'
        st.rerun()


if "user" not in st.session_state or st.session_state["user"] is None:
    st.warning("Veuillez vous connecter d’abord.")
    st.stop()
if st.session_state['user'].get('role') not in {'Admin', 'Responsable'}:
    st.error('Accès réservé aux responsables et administrateurs.')
    st.stop()

if st.session_state.get('_local_preview'):
    api_url, headers = None, {}
else:
    api_url = st.secrets["URL_PRESENCE"]
    headers = {"Authorization": f"Bearer {st.secrets['PRESENCE_TOKEN']}", "Content-Type": "application/json"}
username = st.session_state["user"]["name"]
workspace = load_workspace(api_url, headers, username)
update_home_metrics(workspace)

st.title("Gestion des fiches")
st.caption("Attendance and internship allowance management")
if workspace.get('employes_data') and not workspace.get('legacy_imported'):
    st.info('Une ancienne configuration a été trouvée. Vous pouvez importer ses fiches ; les données sources sont conservées.')
    if st.button('Importer les anciennes fiches'):
        try:
            count = import_legacy_forms(workspace)
            st.success(f'{count} fiche(s) importée(s). Vérifiez les champs avant export.')
        except (ValueError, TypeError, KeyError) as error:
            st.error(f'Import impossible : {error}')

initial_management_tab = st.session_state.pop('management_initial_tab', 'Personnes')
people_tab, create_tab, queue_tab, history_tab, trash_tab = st.tabs(
    ["Personnes", "Créer une fiche", "File d’export", "Historique", "Corbeille"],
    default=initial_management_tab,
)

with people_tab:
    st.subheader("Bibliothèque de personnes")
    with st.form("create_person"):
        person_type = st.radio("Type", ["employee", "intern"], format_func=lambda value: "Salarié" if value == "employee" else "Stagiaire", horizontal=True)
        first_column, second_column = st.columns(2)
        first_name = first_column.text_input("Prénom")
        last_name = second_column.text_input("Nom")
        supervisor_name = first_column.text_input("Responsable")
        supervisor_email = second_column.text_input("E-mail du responsable")
        if st.form_submit_button("Créer la personne", type="primary"):
            if not last_name.strip():
                st.error("Le nom est obligatoire.")
            else:
                workspace["people"].append(create_person(person_type, first_name=first_name, last_name=last_name, supervisor_name=supervisor_name, supervisor_email=supervisor_email))
                st.success("La personne a été ajoutée. Configurez son planning ci-dessous.")

    search_term = st.text_input("Rechercher une personne", key="people_search").casefold()
    people_type = st.selectbox('Type de personne à afficher', ['all', 'employee', 'intern'],
        format_func=lambda value: {'all': 'Tous', 'employee': 'Salarié', 'intern': 'Stagiaire'}[value])
    for person in workspace["people"]:
        if people_type != 'all' and person['type'] != people_type:
            continue
        if search_term and search_term not in person_display_name(person).casefold():
            continue
        with st.expander(f"{person_display_name(person)} — {'Salarié' if person['type'] == 'employee' else 'Stagiaire'}"):
            render_details_editor(person, f"person_{person['id']}")
            render_schedule_editor(person)

with create_tab:
    st.subheader("Nouvelle fiche mensuelle")
    creation_mode = st.radio("Source", ["Personne existante", "Fiche vierge"], horizontal=True, key='creation_source')
    today = datetime.now()
    first_column, second_column = st.columns(2)
    form_month = first_column.number_input("Mois", 1, 12, today.month)
    form_year = second_column.number_input("Année", 2000, 2100, today.year)
    selected_person = None
    blank_type = "employee"
    blank_first_name = ""
    blank_last_name = ""
    if creation_mode == "Personne existante":
        if workspace["people"]:
            person_options = {person["id"]: person for person in workspace["people"]}
            selected_id = st.selectbox("Personne", list(person_options), format_func=lambda item: person_display_name(person_options[item]))
            selected_person = person_options[selected_id]
        else:
            st.info("Créez d’abord une personne ou choisissez une fiche vierge.")
    else:
        blank_type = st.radio("Type de fiche", ["employee", "intern"], format_func=lambda value: "Salarié" if value == "employee" else "Stagiaire", horizontal=True)
        blank_first_name = first_column.text_input("Prénom de la fiche")
        blank_last_name = second_column.text_input("Nom de la fiche")

    if st.button("Ajouter à la file", type="primary"):
        if creation_mode == "Personne existante" and selected_person is None:
            st.error("Sélectionnez une personne existante ou choisissez une fiche vierge.")
            st.stop()
        if creation_mode == "Fiche vierge" and not blank_last_name.strip():
            st.error("Le nom est obligatoire pour créer une fiche vierge.")
            st.stop()
        candidate = create_form(int(form_year), int(form_month), selected_person, person_type=blank_type, first_name=blank_first_name, last_name=blank_last_name)
        duplicate = find_duplicate_form(workspace["forms"], candidate)
        if duplicate:
            st.warning(f"Une fiche existe déjà pour {get_form_name(duplicate)} — {duplicate['month']:02d}/{duplicate['year']}.")
        else:
            workspace["forms"].append(candidate)
            update_home_metrics(workspace)
            st.success("La fiche a été ajoutée à la file.")

with queue_tab:
    st.subheader("File d’export")
    if st.session_state.get('_local_preview'):
        st.caption('Les brouillons non sauvegardés restent dans la session. Utilisez le bouton de sauvegarde sur cet ordinateur pour conserver vos données de test après fermeture du navigateur.')
    else:
        st.caption('Les brouillons de cette version restent dans la session. Pour les retrouver après fermeture du navigateur, utilisez la sauvegarde serveur. Le stockage local durable reste à intégrer.')
    queue_forms = [item for item in workspace["forms"] if not item.get("archived")]
    queue_forms = filter_forms(queue_forms, 'queue')
    if not queue_forms:
        st.info("Aucune fiche en attente.")
    for form in queue_forms:
        with st.expander(f"{get_form_name(form)} — {form['month']:02d}/{form['year']} — {FORM_STATE_LABELS[form['state']]}"):
            render_details_editor(form, f"form_{form['id']}", is_form=True)
            render_calendar_editor(form)
            errors = validate_form(form)
            if errors:
                for error in errors:
                    st.error(error)
            first_column, second_column, third_column = st.columns(3)
            if first_column.button("Marquer terminé", key=f"complete_{form['id']}", disabled=bool(errors)):
                form["state"] = "complete"
                st.rerun()
            if second_column.button("Archiver", key=f"archive_{form['id']}", disabled=form["state"] != "exported" or bool(errors)):
                form["state"] = "archived"
                form["archived"] = True
                update_home_metrics(workspace)
                st.rerun()
            if third_column.button("Supprimer", key=f"trash_{form['id']}"):
                move_to_trash(workspace, form["id"])
                update_home_metrics(workspace)
                st.rerun()
    export_candidates = [item for item in queue_forms if item["state"] in {"complete", "exported"}]
    if export_candidates:
        st.divider()
        st.subheader("Exporter la sélection")
        candidate_options = {item["id"]: item for item in export_candidates}
        selected_ids = st.multiselect(
            "Fiches à exporter",
            list(candidate_options),
            default=list(candidate_options),
            format_func=lambda item: f"{get_form_name(candidate_options[item])} — {candidate_options[item]['month']:02d}/{candidate_options[item]['year']}",
        )
        selected_formats = st.multiselect(
            "Formats", ["word", "pdf", "excel"], default=["pdf"], format_func=lambda item: EXPORT_LABELS[item]
        )
        export_mode = st.radio("Mode", ["Fusionner les fiches", "Créer un ZIP par personne"], horizontal=True)
        selected_forms = [candidate_options[item] for item in selected_ids]
        if st.button("Préparer les exports", type="primary", disabled=not selected_forms or not selected_formats):
            st.session_state.pop(f'exports_{username}', None)
            try:
                st.session_state[f'exports_{username}'] = prepare_exports(selected_forms, selected_formats, export_mode == 'Créer un ZIP par personne')
                st.success('Les fichiers sont prêts. Cliquez sur Télécharger pour les exporter.')
            except Exception as error:
                st.error(f'Export impossible : {error}')
        package = st.session_state.get(f'exports_{username}')
        if package and package['ids'] == selected_ids and package['fingerprint'] == export_fingerprint(selected_forms):
            for file in package['files']:
                st.download_button('Télécharger ' + file['name'], file['data'], file['name'], file['mime'],
                    on_click=mark_downloaded, args=(workspace, package, file['formats']), key=f"download_{username}_{file['name']}")
        if st.button('Archiver la sélection', disabled=not selected_forms or any(item['state'] != 'exported' or validate_form(item) for item in selected_forms)):
            for form in selected_forms:
                form.update(state='archived', archived=True)
            st.rerun()

with history_tab:
    st.subheader("Historique")
    archived_forms = [item for item in workspace["forms"] if item.get("archived")]
    archived_forms = filter_forms(archived_forms, 'history')
    if not archived_forms:
        st.info("Aucune fiche archivée.")
    for form in archived_forms:
        with st.expander(f"Consulter / corriger : {get_form_name(form)} — {form['month']:02d}/{form['year']}"):
            if st.checkbox('Modifier cette fiche archivée', key=f"edit_archive_{form['id']}"):
                render_details_editor(form, f"archived_{form['id']}", is_form=True)
                render_calendar_editor(form)
            format_name = st.selectbox('Format de réexport', ['word', 'pdf', 'excel'], format_func=lambda value: EXPORT_LABELS[value], key=f"history_format_{form['id']}")
            if st.button('Préparer le fichier', key=f"history_export_{form['id']}"):
                try:
                    st.session_state[f"history_export_{username}_{form['id']}"] = prepare_exports([form], [format_name])
                except Exception as error:
                    st.error(f'Export impossible : {error}')
            history_package = st.session_state.get(f"history_export_{username}_{form['id']}")
            if history_package and history_package['fingerprint'] == export_fingerprint([form]):
                file = history_package['files'][0]
                st.download_button('Télécharger', file['data'], file['name'], file['mime'], key=f"history_download_{form['id']}")
        columns = st.columns([4, 1, 1])
        columns[0].write(f"**{get_form_name(form)}** — {form['month']:02d}/{form['year']}")
        if columns[1].button("Désarchiver", key=f"unarchive_{form['id']}"):
            form["archived"] = False
            form["state"] = "exported"
            update_home_metrics(workspace)
            st.rerun()
        if columns[2].button("Corbeille", key=f"history_trash_{form['id']}"):
            move_to_trash(workspace, form["id"])
            update_home_metrics(workspace)
            st.rerun()

with trash_tab:
    st.subheader("Corbeille")
    if not workspace["trash"]:
        st.info("La corbeille est vide.")
    for form in workspace["trash"]:
        columns = st.columns([4, 1, 1])
        columns[0].write(f"**{get_form_name(form)}** — supprimée le {form.get('deleted_at', '')[:10]}")
        if columns[1].button("Restaurer", key=f"restore_{form['id']}"):
            try:
                restore_from_trash(workspace, form["id"])
                update_home_metrics(workspace)
                st.rerun()
            except ValueError as error:
                st.error(str(error))
        if columns[2].button('Supprimer définitivement', key=f"permanent_{form['id']}"):
            st.session_state[f'pending_delete_{username}'] = form['id']

if st.session_state.get(f'pending_delete_{username}'):
    confirm_permanent_delete(username, st.session_state[f'pending_delete_{username}'])

st.divider()
save_label = 'Sauvegarder les données de test sur cet ordinateur' if st.session_state.get('_local_preview') else 'Sauvegarder les données sur le serveur'
if st.button(save_label, use_container_width=True):
    save_workspace(api_url, headers, username, workspace)
