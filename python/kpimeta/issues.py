"""Problems found while reading, composing, validating or writing data."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

ERROR = "error"
WARNING = "warning"
INFO = "info"

_ICONS = {ERROR: "ERRORE", WARNING: "AVVISO", INFO: "INFO"}


@dataclass
class Issue:
    severity: str
    message: str
    row: int | None = None  # 0-based row of the batch, None = whole batch/DB
    column: str | None = None
    value: Any = None

    def label(self) -> str:
        where = []
        if self.row is not None:
            where.append(f"riga {self.row + 1}")
        if self.column:
            where.append(self.column)
        prefix = f"[{', '.join(where)}] " if where else ""
        return f"{_ICONS.get(self.severity, self.severity)}: {prefix}{self.message}"


def errors(issues: list[Issue]) -> list[Issue]:
    return [i for i in issues if i.severity == ERROR]


def warnings(issues: list[Issue]) -> list[Issue]:
    return [i for i in issues if i.severity == WARNING]
