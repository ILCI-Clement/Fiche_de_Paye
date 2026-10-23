"""Required structured names for new login accounts; no inference from old labels."""
from pydantic import BaseModel, ConfigDict, Field, model_validator


class AccountNames(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True, hide_input_in_errors=True)
    first_name: str = Field(strict=True, min_length=1, max_length=100)
    last_name: str = Field(strict=True, min_length=1, max_length=100)

    @model_validator(mode='after')
    def normalized(self):
        self.first_name = ' '.join(self.first_name.split())
        self.last_name = ' '.join(self.last_name.split())
        if not self.first_name or not self.last_name or len(self.full_name) > 200:
            raise ValueError('Renseignez les prénoms et le nom (200 caractères maximum au total).')
        if any(ord(character) < 32 for character in self.full_name):
            raise ValueError('Les prénoms et le nom contiennent un caractère invalide.')
        return self

    @property
    def full_name(self):
        return f'{self.first_name} {self.last_name}'


def ensure_account_names_schema(cursor, columns):
    additions = [f"ADD COLUMN {key} VARCHAR(100) NOT NULL DEFAULT ''"
                 for key in ('first_name', 'last_name') if key not in columns]
    if not additions:
        return
    # Old users tables may retain zero-date defaults. Keep this compatibility
    # local to the additive ALTER and restore the original session SQL mode.
    cursor.execute('SELECT @@SESSION.sql_mode AS sql_mode')
    mode = cursor.fetchone()['sql_mode']
    try:
        cursor.execute('SET SESSION sql_mode=%s', (','.join(
            item for item in mode.split(',') if item not in {'NO_ZERO_DATE', 'NO_ZERO_IN_DATE'}),))
        cursor.execute('ALTER TABLE users ' + ','.join(additions))
    finally:
        cursor.execute('SET SESSION sql_mode=%s', (mode,))
