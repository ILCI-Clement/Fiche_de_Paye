# Dev Requirements — Presence

## 1. Objective

The application manages monthly presence sheets for employees and interns. It must allow a manager to prepare, review, save, export and archive one or several sheets while preserving unfinished work.

The data contract remains compatible with the current Presence API: the monthly form is stored as JSON in the `Presence.form_content` field of MariaDB.

## 2. Access and roles

- Keep the existing login mechanism and the current role hierarchy (`Admin`, `Responsable`, `Employe`).
- A user can see and manage only the forms and personnel records created by that user.
- Personnel management is separate from form creation; a form must not silently create or overwrite a personnel record.
- The rule defining who enters the daily presence data (manager only or employee participation) remains a business decision to confirm. The UI must not hard-code that decision prematurely.

## 3. Personnel and form creation

- The home page focuses on creating a new presence form.
- A new form can be created from an existing employee/intern record or as a blank form.
- Employees and interns both have a presence form.
- A person already added to the queue must never be duplicated automatically.
- Removing an unfinished form is an explicit user action; partially completed data must be preserved until the user chooses to delete it.

### Default schedule

- Each saved employee or intern may have an individual default schedule (for example, 09:00–17:00 or 09:30–17:30).
- The standard default schedule is Monday through Friday; Saturday and Sunday are non-working days unless the responsible person explicitly enables them for that person.
- Selecting a person for a new form pre-fills that schedule.
- The schedule can be edited manually for the current form.
- Changing a person's default schedule affects new forms only. Existing, unfinished or archived forms must not be updated automatically.
- Monthly leave, paid leave (`congé payé`), public holidays and other exceptions remain manually editable on the form.

## 4. Queue and persistence

- The user can add several forms to an export queue and continue editing them later.
- Saving to the server is an explicit choice by the person filling the form.
- Unfinished forms remain in the queue until the user removes them.
- Deletion first moves a form to a recycle bin; permanent deletion requires a second confirmation.
- Deleted forms can be restored. Items in the recycle bin are permanently removed after 60 days.
- Queue operations are manual; no automatic deletion or automatic re-creation is allowed.

## 5. Monthly visual calendar

The presence form must offer a clear monthly calendar view inspired by Google Calendar.

- Display the selected month in a conventional calendar grid, with working days arranged by week.
- Each day is divided into morning and afternoon segments so that half-day values can be entered independently.
- The default calendar is a read-only overview. A `Modifier` action explicitly opens batch-edit mode.
- In batch-edit mode, each morning and afternoon is an independent, fixed-size cell. The user can click or click-drag across several half-days before opening one contextual status menu for the whole selection.
- A drag that starts on an unselected half-day adds the crossed half-days to the selection; a drag that starts on a selected half-day removes them. Every half-day card uses the same fixed height; longer labels are truncated instead of resizing the grid.
- Selection is local to the calendar until a status is confirmed. Selecting, deselecting or dragging must not trigger a Streamlit page rerun; the form refreshes only when the user applies a status or leaves edit mode.
- Applying a status automatically clears the selection. `Terminer` exits edit mode and `Annuler la sélection` clears a pending selection without changing the calendar.
- The optional `Autre` reason is entered and confirmed in the batch status menu; no secondary editor is displayed below the calendar.
- Choosing `Travail` restores the person's default schedule for that half-day. Per-half-day time editing is out of scope for this version.
- The grid uses the application's dark visual language. Statuses use readable, low-saturation accent colors without a light calendar background.
- Calendar columns retain a compact, consistent width. Every half-day card uses the same taller fixed height so default times can wrap and remain visible without changing individual card dimensions.
- Support at least: working day, paid leave, absence, sick leave, public holiday and `Autre` with a manually entered label.
- Consecutive working days with the same schedule are rendered as one continuous visual bar. The first and last segment of the sequence have rounded ends; intermediate segments remain visually connected. A break in the sequence starts a new bar.
- The calendar is a visual editor for the existing data model. It must preserve the current `planning_detail`, `vacances`, `absences` and `arret` semantics and must not silently alter archived forms.
- Editing a calendar cell must immediately update the form state and remain compatible with manual corrections before saving or exporting.

### Calendar acceptance criteria

1. Opening a month shows all dates and both half-days without leaving the form page.
2. Batch-edit mode allows several independently selected half-days, including by click-drag selection, to receive the same status from one contextual menu.
3. A sequence of consecutive working days is displayed as a connected bar with rounded ends.
4. A sequence interrupted by a leave, absence, holiday or non-working day is split into separate bars.
5. The contextual menu applies its selected status and custom `Autre` text only to the explicitly selected half-days.
6. Existing saved JSON can be loaded into the calendar without losing fields that are not displayed by the calendar.

## 6. Export and archive

- Offer individual export in Word, PDF or Excel, with one format selected at a time.
- Support combined export: each person gets a separate page and a separate table in the combined file.
- Support batch separate export: the archive contains one folder per person, named after that person, with that person's file(s).
- Both combined and separate batch export modes must be available.
- Individual filenames include the person's name and the month/year prefix.
- After export, the manager chooses whether to archive the result, archive several results in batch or leave them in the queue.
- Unarchived exports remain visible in the queue.
- Archived exports are frozen and do not follow later personnel or default-schedule changes automatically.
- An explicit editable option may be used for an exceptional post-archive correction. A form can be marked as unarchived and returned to the queue; it does not need to be archived again automatically.

## 7. Employee and intern data rules

- Employee salary and payroll amounts are not part of the presence form.
- Intern compensation fields required by the existing internship document (allowance, hourly rate, transport and related values) are retained.
- Validation must prevent incomplete dates, invalid time ranges and contradictory half-day selections before generation.

## 8. Language and implementation constraints

- The user interface is primarily French, with English as a secondary language where useful.
- Source code identifiers, code comments and technical names are written in English; Chinese text must not be introduced into the project files.
- Preserve the current authentication flow and the MariaDB/Presence API data contract unless an explicit migration is designed and tested.
- Password recovery links must open a dedicated password-reset view. The view validates the password confirmation locally, submits the URL token and new password to the Presence API, handles expired or invalid links without exposing account information, and clears the token from the browser URL after a successful reset.
- Every change to persistence, generation or export must be covered by a regression test before production deployment.
