"""
The reference data, read out of the workbook once per run.

All of it is found by table or defined name, never by sheet and column.
The project data has already been split across sheets twice, and every
hardcoded column reference broke silently when it happened.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from . import config
from .text import area_key, tokenize, match_key
from .workbook import LiveWorkbook, text_of


@dataclass
class ReferenceData:
    # activity -> the activities that must come before it
    previous: dict[str, set[str]] = field(default_factory=dict)

    # parent activity -> its sub-activities, and the reverse
    split_subs: dict[str, set[str]] = field(default_factory=dict)
    split_parents: dict[str, set[str]] = field(default_factory=dict)

    # "P08-A1-008-GF" -> the apartments that floor really has
    apartments: dict[str, set[str]] = field(default_factory=dict)

    # status code -> (meaning, counts as approved)
    status: dict[str, tuple[str, bool]] = field(default_factory=dict)

    # every accepted spelling -> canonical name, plus the canonical list
    area_canonical: dict[str, str] = field(default_factory=dict)
    area_names: list[str] = field(default_factory=list)
    area_floors: dict[str, set[str]] = field(default_factory=dict)

    # -- lookups ------------------------------------------------------------

    def status_text(self, code: str) -> str:
        code = (code or "").strip()
        if not code:
            return "No status"
        meaning = self.status.get(code, ("", False))[0]
        return meaning or f"Status {code}"

    def is_approved(self, code: str) -> bool:
        code = (code or "").strip()
        if not code:
            return False
        if code in self.status:
            return self.status[code][1]
        return code.upper() == config.DEFAULT_APPROVED_CODE

    def canonical_area(self, token: str) -> str | None:
        return self.area_canonical.get(area_key(token))

    def floor_allowed(self, canonical: str, floor: str) -> bool:
        """A blank Floors cell means anywhere.

        Main Electrical Room and Main Telephone Room are basement-only and
        are NOT the same rooms as Electrical Room / Telephone Room -
        collapsing them into one family, as generic room matching tends to,
        would hide a real mistake.
        """
        allowed = self.area_floors.get(canonical)
        if not allowed or not floor.strip():
            return True
        return floor.strip().upper() in allowed


def load(book: LiveWorkbook) -> ReferenceData:
    data = ReferenceData()
    _load_previous(book, data)
    _load_splits(book, data)
    _load_apartments(book, data)
    _load_status(book, data)
    _load_areas(book, data)
    return data


def _load_previous(book: LiveWorkbook, data: ReferenceData) -> None:
    for row in book.table_values(config.TABLE_PREV_ACTIVITIES):
        current = text_of(row[0] if row else "")
        if not current:
            continue
        bucket = data.previous.setdefault(current, set())
        for cell in row[1:]:
            previous = text_of(cell)
            if previous:
                bucket.add(previous)


def _load_splits(book: LiveWorkbook, data: ReferenceData) -> None:
    """Activities logged as two or more sub-activities instead of one.
    Wires_Pulling is raised as Wires_H.L + Wire_PWR_Pulling. When every sub
    is approved the parent counts as done, and no WIR for the parent itself
    will ever exist."""
    rows = book.named_values(config.NAME_ACTIVITY_SPLITS)
    for row in rows[1:]:                       # row 1 is the header
        parent = text_of(row[0] if row else "")
        if not parent:
            continue
        subs = data.split_subs.setdefault(parent, set())
        for cell in row[1:]:
            child = text_of(cell)
            if child:
                subs.add(child)
                data.split_parents.setdefault(child, set()).add(parent)


def _load_apartments(book: LiveWorkbook, data: ReferenceData) -> None:
    """Prefers the Apt01..Apt15 columns - one unit per cell, nothing to
    parse - and falls back to the free-text Apartment column."""
    for row in book.table_values(config.TABLE_FLOOR_APARTMENTS):
        code = text_of(row[config.REG_CODE] if row else "").upper()
        if not code or code in data.apartments:
            continue

        units = set()
        for index in range(config.REG_APT_FIRST, min(config.REG_APT_LAST + 1, len(row))):
            unit = text_of(row[index])
            if unit:
                units.add(match_key(unit))

        if not units and len(row) > config.REG_APT_TEXT:
            for token in tokenize(text_of(row[config.REG_APT_TEXT])):
                key = match_key(token)
                if key:
                    units.add(key)

        data.apartments[code] = units


def _load_status(book: LiveWorkbook, data: ReferenceData) -> None:
    """Code | Meaning | Counts as approved. Falls back to 'B means
    approved' if the table is missing, which is what the code did before
    the table existed."""
    rows = book.named_values(config.NAME_STATUS_CODES)
    for row in rows[1:]:
        code = text_of(row[0] if row else "")
        if not code or code in data.status:
            continue
        meaning = text_of(row[1]) if len(row) > 1 else ""
        approved_cell = text_of(row[2]) if len(row) > 2 else ""
        data.status[code] = (meaning, approved_cell[:1].upper() == "Y")

    if not data.status:
        data.status[config.DEFAULT_APPROVED_CODE] = ("Approved", True)


def _load_areas(book: LiveWorkbook, data: ReferenceData) -> None:
    """Area | Floors | Variant 1..n. The first column is how the area
    should be written; the rest are spellings seen in the log that mean the
    same place."""
    rows = book.named_values(config.NAME_AREA_NAMES)
    for row in rows[1:]:
        canonical = text_of(row[0] if row else "")
        if not canonical:
            continue

        if canonical not in data.area_floors:
            data.area_names.append(canonical)
            floors = text_of(row[1]) if len(row) > 1 else ""
            data.area_floors[canonical] = {
                f.strip().upper() for f in floors.split(",") if f.strip()
            }

        data.area_canonical.setdefault(area_key(canonical), canonical)
        for cell in row[2:]:
            variant = text_of(cell)
            if variant:
                data.area_canonical.setdefault(area_key(variant), canonical)
