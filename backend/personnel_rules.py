"""Validation and stable presentation rules for the independent personnel register."""
from datetime import date, datetime
from typing import Literal
from zoneinfo import ZoneInfo

from pydantic import BaseModel, ConfigDict, Field, SecretStr, StrictBool, field_validator, model_validator

try:
    from .personnel_sensitive import NirProtectionError, normalize_social_security_number
except ImportError:
    from personnel_sensitive import NirProtectionError, normalize_social_security_number


PersonnelType = Literal['', 'salarie', 'stagiaire', 'alternance']
ContractType = Literal['', 'CDI', 'CDD', 'stage', 'autre']
ArchiveFilter = Literal['active', 'archived', 'all']
EmploymentStatus = Literal['employed', 'left', 'upcoming', 'unknown', 'all']
MAX_IDENTIFIER = 9_007_199_254_740_991


def paris_today():
    return datetime.now(ZoneInfo('Europe/Paris')).date()


class RegisterInput(BaseModel):
    model_config = ConfigDict(extra='forbid', str_strip_whitespace=True, hide_input_in_errors=True)


class PeriodInput(RegisterInput):
    entry_date: date | None = None
    # Legacy exit_date remains the account contract end, never an actual departure.
    exit_date: date | None = None
    actual_exit_date: date | None = None
    contract_type: ContractType = ''
    other_contract: str = Field(default='', max_length=200)
    is_foreign_worker: StrictBool = False
    work_permit_type: str = Field(default='', max_length=200)
    work_permit_number: str = Field(default='', max_length=200)
    is_temporary_worker: StrictBool = False
    temporary_company_name: str = Field(default='', max_length=200)
    temporary_company_address: str = Field(default='', max_length=1000)
    is_young_worker: StrictBool = False
    youth_contract_type: Literal['', 'apprenti', 'qualification', 'adaptation'] = ''

    @model_validator(mode='after')
    def check_dates(self):
        if self.entry_date and self.exit_date and self.exit_date < self.entry_date:
            raise ValueError('La fin prévue du contrat ne peut pas précéder l’entrée.')
        if self.entry_date and self.actual_exit_date and self.actual_exit_date < self.entry_date:
            raise ValueError('La sortie effective ne peut pas précéder l’entrée.')
        if self.is_foreign_worker:
            if not self.work_permit_type or not self.work_permit_number:
                raise ValueError('Renseignez le type et le numéro du document autorisant le travail.')
        if self.is_temporary_worker:
            if not self.contract_type or (self.contract_type == 'autre' and not self.other_contract):
                raise ValueError('Renseignez le type de contrat du travailleur temporaire.')
            if not self.temporary_company_name or not self.temporary_company_address:
                raise ValueError('Renseignez le nom et l’adresse de l’entreprise de travail temporaire.')
        if self.is_young_worker and not self.youth_contract_type:
            raise ValueError('Sélectionnez le contrat du jeune travailleur ou de l’apprenti.')
        return self

    def storage_values(self):
        values = self.model_dump(include=set(PeriodInput.model_fields))
        for flag, fields in (
            ('is_foreign_worker', ('work_permit_type', 'work_permit_number')),
            ('is_temporary_worker', ('temporary_company_name', 'temporary_company_address')),
            ('is_young_worker', ('youth_contract_type',)),
        ):
            if not values[flag]:
                for key in fields:
                    values[key] = ''
        return values


class EmployeeInput(RegisterInput):
    full_name: str = Field(default='', max_length=200)
    last_name: str = Field(default='', max_length=100)
    first_name: str = Field(default='', max_length=100)
    nationality: str = Field(default='', max_length=100)
    sex: Literal['', 'F', 'M'] = ''
    address: str = Field(default='', max_length=1000)
    birth_date: date | None = None
    job_title: str = Field(default='', max_length=200)
    qualification: str = Field(default='', max_length=200)
    personnel_type: PersonnelType = ''
    linked_username: str | None = Field(default=None, max_length=255)
    # Exclusion prevents ordinary values, JSON dumps and audit snapshots collecting the NIR.
    social_security_number: SecretStr | None = Field(default=None, exclude=True, repr=False)

    @model_validator(mode='after')
    def display_name(self):
        if self.first_name and self.last_name:
            name = f'{self.first_name} {self.last_name}'
            if len(name) > 200:
                raise ValueError('Le nom et les prénoms sont limités à 200 caractères au total.')
            self.full_name = name
        if not self.full_name:
            raise ValueError('Renseignez le nom et les prénoms.')
        return self

    @model_validator(mode='before')
    @classmethod
    def protect_sensitive_validation_input(cls, value):
        if isinstance(value, dict) and 'social_security_number' in value:
            value = dict(value)
            raw = value['social_security_number']
            if raw is not None and not isinstance(raw, SecretStr):
                # Wrap before validation: even ValidationError.errors() holds a masked
                # SecretStr rather than the original number. Non-string input stays invalid.
                value['social_security_number'] = SecretStr(raw if isinstance(raw, str) else 'invalid-input')
        return value

    @field_validator('social_security_number')
    @classmethod
    def check_social_security_number(cls, value):
        if value is None:
            return None
        try:
            normalized = normalize_social_security_number(value.get_secret_value())
        except NirProtectionError:
            raise ValueError('Le numéro de sécurité sociale doit respecter le format NIR de 15 caractères.') from None
        return SecretStr(normalized) if normalized is not None else None

    @field_validator('birth_date')
    @classmethod
    def check_birth_date(cls, value):
        if value and value > paris_today():
            raise ValueError('La date de naissance ne peut pas être future.')
        return value

    @field_validator('linked_username')
    @classmethod
    def blank_link_is_none(cls, value):
        return value or None


class EmployeeCreate(EmployeeInput):
    initial_period: PeriodInput | None = None

    @model_validator(mode='after')
    def required_registration(self):
        missing = registration_missing(self.model_dump(), self.initial_period.model_dump() if self.initial_period else None)
        if missing:
            raise ValueError('Champs obligatoires : ' + ', '.join(missing) + '.')
        if self.birth_date and self.initial_period.entry_date < self.birth_date:
            raise ValueError('La date d’entrée ne peut pas précéder la naissance.')
        return self


class EmployeeUpdate(EmployeeInput):
    expected_version: int = Field(strict=True, ge=1, le=2_147_483_646)


class PeriodMutation(PeriodInput):
    expected_version: int = Field(strict=True, ge=1, le=2_147_483_646)


class ArchiveMutation(RegisterInput):
    is_archived: StrictBool
    expected_version: int = Field(strict=True, ge=1, le=2_147_483_646)


def employment_status(period, today=None):
    """Missing entry dates remain unknown; an explicit entry with no exit is open-ended."""
    if not period or not period.get('entry_date'):
        return 'unknown'
    today = today or paris_today()
    entry = period['entry_date']
    end = period.get('actual_exit_date')
    if isinstance(entry, str):
        entry = date.fromisoformat(entry)
    if isinstance(end, str):
        end = date.fromisoformat(end)
    if entry > today:
        return 'upcoming'
    if end and end < today:
        return 'left'
    return 'employed'


def registration_missing(employee, period=None):
    """Legacy records remain editable; completeness never blocks unrelated operations."""
    labels = {'last_name': 'Nom', 'first_name': 'Prénoms', 'nationality': 'Nationalité',
              'sex': 'Sexe', 'birth_date': 'Date de naissance', 'address': 'Adresse',
              'job_title': 'Emploi', 'qualification': 'Qualification', 'personnel_type': 'Type de personnel'}
    missing = [label for key, label in labels.items() if not employee.get(key)]
    if not period or not period.get('entry_date'):
        missing.append('Date d’entrée')
    return missing


def like_search(value):
    """Use '=' as the LIKE escape so user wildcards do not broaden the search."""
    return '%' + value.replace('=', '==').replace('%', '=%').replace('_', '=_') + '%'
