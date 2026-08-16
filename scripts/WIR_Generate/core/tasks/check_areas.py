"""
Check Areas - writes into column K and nowhere else.

The Area column itself is never overwritten, so nothing is lost if a
suggestion is wrong: K gets the suggested Area text when the value can be
tidied, and the reason when it cannot. There is no report sheet - the
answer belongs next to the row it is about.

Floors without a defined area scheme are skipped entirely.
"""

from __future__ import annotations

from .. import config
from ..findings import RunResult
from ..reference import ReferenceData
from ..text import (
    is_unit_code,
    match_key,
    nearest,
    normalize_spaces,
    split_glued_numbers,
    tokenize,
)
from ..workbook import LiveWorkbook, text_of

TITLE = "Check Areas"


def floor_is_checkable(floor: str) -> bool:
    return floor.strip().upper() in config.CHECKED_FLOORS


def format_area(area: str, floor: str, ref: ReferenceData) -> str:
    """Rewrites one Area value into the house format:

        "301 302 303 304"        -> "Apartments 301, 302, 303, 304"
        "101,102,103"            -> "Apartments 101, 102, 103"
        "Cooridor"               -> "Corridor"
        "MTR"                    -> whatever Area_Names calls it

    Unit codes are collected under one "Apartments" prefix in the order
    they were typed - not sorted, because the typed order sometimes
    carries meaning and reordering it silently would be surprising.
    Anything unknown is left exactly as written.
    """
    if not area.strip():
        return area

    units: list[str] = []
    names: list[str] = []
    seen: set[str] = set()

    for token in tokenize(area):
        key = match_key(token)
        if not key:
            continue

        if is_unit_code(key):
            # Recover a comma list Excel turned into one long number.
            for unit in split_glued_numbers(key, floor) or [key]:
                if unit not in seen:
                    seen.add(unit)
                    units.append(unit)
        elif key.startswith("PART"):
            if token not in seen:
                seen.add(token)
                names.append(token)
        else:
            canonical = ref.canonical_area(token) or token
            if canonical not in seen:
                seen.add(canonical)
                names.append(canonical)

    parts: list[str] = []
    if units:
        parts.append("Apartments " + ", ".join(units))
    if names:
        parts.append(", ".join(names))

    return ", ".join(parts) if parts else area


def area_comment(area: str, floor: str, package: str, building: str,
                 ref: ReferenceData) -> str:
    """Everything about this row's Area that could not be resolved."""
    if not area.strip():
        return ""

    code = f"{package}-{building}-{floor}".upper()
    known = ref.apartments.get(code, set())
    notes: list[str] = []

    for token in tokenize(area):
        key = match_key(token)
        if not key:
            continue

        if is_unit_code(key):
            if known and key not in known and not split_glued_numbers(key, floor):
                notes.append(f"{key} is not in the register for {code}.")
            continue

        if key.startswith("PART"):
            continue

        canonical = ref.canonical_area(token)
        if canonical:
            if not ref.floor_allowed(canonical, floor):
                allowed = ", ".join(sorted(ref.area_floors[canonical]))
                notes.append(f"{canonical} is only on {allowed}, not {floor}.")
        else:
            suggestion = nearest(token, ref.area_names)
            if suggestion:
                notes.append(f"{token} is not a known area - did you mean {suggestion}?")
            else:
                notes.append(
                    f"{token} is not a known area. "
                    f"Add it to {config.NAME_AREA_NAMES} if it is correct."
                )

    return " ".join(notes)


def run(book: LiveWorkbook, ref: ReferenceData) -> RunResult:
    result = RunResult(TITLE)
    sheet = book.log_sheet

    last_row = book.last_row(sheet, config.Main.WIR_NO)
    if last_row < config.FIRST_DATA_ROW:
        return result.fail(f"No rows found in column B of '{sheet.Name}'.")

    rows = book.block(sheet, config.FIRST_DATA_ROW, last_row, config.Main.LAST_COL)

    # Last run's output goes first, or a fixed row keeps its old note.
    count = last_row - config.FIRST_DATA_ROW + 1
    book.write_column(sheet, config.Main.SUGGEST, config.FIRST_DATA_ROW, [""] * count)
    book.clear_fill(sheet, config.Main.SUGGEST, config.FIRST_DATA_ROW, last_row)

    column: list[str] = []
    checked = suggested = commented = skipped_floor = skipped_done = 0
    flag_rows: list[int] = []

    for offset, values in enumerate(rows):
        row_no = config.FIRST_DATA_ROW + offset
        column.append("")

        wir_no = text_of(values[config.Main.WIR_NO - 1])
        if not wir_no:
            continue

        if text_of(values[config.Main.STATUS - 1]).lower() == config.STATUS_DONE.lower():
            skipped_done += 1
            continue

        floor = text_of(values[config.Main.FLOOR - 1])
        if not floor_is_checkable(floor):
            skipped_floor += 1
            continue

        checked += 1
        area = normalize_spaces(text_of(values[config.Main.AREA - 1]))
        package = text_of(values[config.Main.PACKAGE - 1])
        building = text_of(values[config.Main.BUILDING - 1])

        suggestion = format_area(area, floor, ref) if area else ""
        if suggestion == area:
            suggestion = ""

        comment = area_comment(area, floor, package, building, ref)

        cell = suggestion
        if comment:
            cell = f"{cell}\n{comment}" if cell else comment

        column[-1] = cell
        if comment:
            commented += 1
            flag_rows.append(row_no)
        elif cell:
            suggested += 1

    book.write_column(sheet, config.Main.SUGGEST, config.FIRST_DATA_ROW, column)
    for row_no in flag_rows:
        book.fill(sheet, row_no, config.Main.SUGGEST, config.COLOUR_WARN)

    sheet.Columns(config.Main.SUGGEST).WrapText = True
    sheet.Columns(config.Main.SUGGEST).VerticalAlignment = -4160   # xlTop

    result.line(f"Rows checked:          {checked}")
    result.line(f"Suggested corrections: {suggested}")
    result.line(f"Needs a look:          {commented}")
    result.line(f"Skipped (floor):       {skipped_floor}")
    result.line(f"Skipped (Done):        {skipped_done}")
    result.line("")
    result.line("Results are in column K. The Area column is untouched.")
    return result
