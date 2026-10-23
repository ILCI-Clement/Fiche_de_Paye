"""Synthetic complete registrations and disposable SQLite schema additions."""


def registration_payload(**overrides):
    values = dict(full_name='Synthetic personnel fixture', first_name='Synthetic', last_name='Fixture',
                  nationality='Synthetic', sex='F', birth_date='2000-01-02', address='Synthetic address',
                  job_title='Synthetic job', qualification='Synthetic qualification', personnel_type='salarie',
                  initial_period={'entry_date': '2026-01-01'})
    if 'full_name' in overrides and 'first_name' not in overrides and 'last_name' not in overrides:
        values.update(first_name=overrides['full_name'], last_name='Fixture')
    return {**values, **overrides}


def extend_sqlite_register(db):
    # MySQL migration is exercised independently with TEMPORARY tables.
    for table, fields in (
        ('personnel_register', {'first_name': "TEXT DEFAULT ''", 'last_name': "TEXT DEFAULT ''",
                                'registration_schema_version': 'INTEGER DEFAULT 1'}),
        ('personnel_register_periods', {'actual_exit_date': 'TEXT', 'is_foreign_worker': 'INTEGER DEFAULT 0',
            'work_permit_type': "TEXT DEFAULT ''", 'work_permit_number': "TEXT DEFAULT ''",
            'is_temporary_worker': 'INTEGER DEFAULT 0', 'temporary_company_name': "TEXT DEFAULT ''",
            'temporary_company_address': "TEXT DEFAULT ''", 'is_young_worker': 'INTEGER DEFAULT 0',
            'youth_contract_type': "TEXT DEFAULT ''"}),
    ):
        for field, definition in fields.items():
            db.execute(f'ALTER TABLE {table} ADD COLUMN {field} {definition}')
