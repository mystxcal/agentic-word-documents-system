from __future__ import annotations

from dataclasses import asdict, dataclass
from enum import StrEnum
from typing import Iterable


class Severity(StrEnum):
    INFO = "info"
    WARNING = "warning"
    ERROR = "error"


@dataclass(frozen=True)
class Diagnostic:
    code: str
    severity: Severity
    message: str
    location: str | None = None
    hint: str | None = None

    def to_dict(self) -> dict:
        return asdict(self)


class DiagnosticBag:
    def __init__(self, values: Iterable[Diagnostic] | None = None):
        self._values = list(values or [])

    def add(
        self,
        code: str,
        severity: Severity,
        message: str,
        *,
        location: str | None = None,
        hint: str | None = None,
    ) -> None:
        self._values.append(Diagnostic(code, severity, message, location, hint))

    def info(self, code: str, message: str, **kwargs) -> None:
        self.add(code, Severity.INFO, message, **kwargs)

    def warn(self, code: str, message: str, **kwargs) -> None:
        self.add(code, Severity.WARNING, message, **kwargs)

    def error(self, code: str, message: str, **kwargs) -> None:
        self.add(code, Severity.ERROR, message, **kwargs)

    @property
    def values(self) -> tuple[Diagnostic, ...]:
        return tuple(self._values)

    @property
    def has_errors(self) -> bool:
        return any(item.severity == Severity.ERROR for item in self._values)

    def to_list(self) -> list[dict]:
        return [item.to_dict() for item in self._values]
