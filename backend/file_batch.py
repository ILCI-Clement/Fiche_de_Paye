"""Shared bounded file actions and post-commit filesystem cleanup."""
from typing import Literal
from pydantic import BaseModel, ConfigDict, Field, StrictBool

DELETE_WORD = 'SUPPRIMER'


class FileBatchAction(BaseModel):
    model_config = ConfigDict(extra='forbid')
    action: Literal['archive', 'restore', 'delete']
    expected_archived: StrictBool
    confirmation: str | None = Field(default=None, strict=True, max_length=32)


def cleanup_files(paths):
    """Never unlink before the transaction/audit commits; retain explicit failure state."""
    pending = False
    for path in paths:
        try:
            path.unlink(missing_ok=True)
        except OSError:
            pending = True
    return pending
