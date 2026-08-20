"""
What a check produces, and what a run reports back.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import IntEnum


class Severity(IntEnum):
    WARN = 1
    ERROR = 2

    @property
    def label(self) -> str:
        return "ERROR" if self is Severity.ERROR else "WARN"


@dataclass
class Finding:
    row: int
    wir_no: str
    severity: Severity
    check: str
    message: str


@dataclass
class RunResult:
    """One task's outcome, in the shape the UI needs to render it."""

    title: str
    ok: bool = True
    headline: str = ""
    details: list[str] = field(default_factory=list)
    findings: list[Finding] = field(default_factory=list)

    @property
    def errors(self) -> int:
        return sum(1 for f in self.findings if f.severity is Severity.ERROR)

    @property
    def warnings(self) -> int:
        return sum(1 for f in self.findings if f.severity is Severity.WARN)

    def line(self, text: str) -> None:
        self.details.append(text)

    def fail(self, message: str) -> "RunResult":
        self.ok = False
        self.headline = message
        return self


# Checks whose subject is the row's own Area, so the Area cell is coloured
# as well as the WIR number.
AREA_CHECKS = {"Apartment", "Area name", "Area submitted before"}
