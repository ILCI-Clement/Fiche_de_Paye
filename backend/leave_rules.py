"""Paid leave input and authorization, independent of HTTP and storage."""
from datetime import date
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field, model_validator

Status = Literal['pending', 'approved', 'rejected', 'withdrawn']
Slot = Literal['full', 'morning', 'afternoon']


class LeaveInput(BaseModel):
    model_config = ConfigDict(extra='forbid')
    start_date: date
    end_date: date
    start_slot: Slot = 'full'
    end_slot: Slot = 'full'
    note: str = Field(default='', max_length=2000)

    @model_validator(mode='after')
    def valid_period(self):
        if self.end_date < self.start_date:
            raise ValueError('La fin doit être postérieure ou égale au début.')
        if self.start_date == self.end_date and self.start_slot != self.end_slot:
            raise ValueError('Pour un seul jour, choisissez la même période au début et à la fin.')
        self.note = self.note.strip()
        return self


class LeaveDecision(BaseModel):
    model_config = ConfigDict(extra='forbid')
    decision: Literal['approved', 'rejected']
    comment: str = Field(default='', max_length=2000)


def roles(actor):
    return set(actor.get('role_tags') or [actor.get('role')])


def can_review(actor, employee_username, current_manager):
    # Department leadership does not replace the direct-manager relationship.
    return 'Admin' in roles(actor) or (
        'Responsable' in roles(actor)
        and actor['username'] != employee_username
        and actor['username'] == current_manager
    )


def display_name(user):
    return str(user.get('display_name') or user.get('email') or user['username']).split('@', 1)[0]
