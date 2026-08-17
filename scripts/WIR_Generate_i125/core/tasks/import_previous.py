"""
Import Previous WIR Paths.

For every pending log row, finds the approved WIRs for the activities that
must come before this one at the same Package / Building / Floor, and
writes their file paths into column AA.

The split rules apply: a parent activity is satisfied by all of its
sub-activities, and a sub by its parent.
"""

from __future__ import annotations

from collections import defaultdict

from .. import config
from ..findings import RunResult
from ..reference import ReferenceData
from ..text import is_unit_code, match_key, tokenize
from ..workbook import LiveWorkbook, text_of

TITLE = "Import Previous WIR Paths"

PATH_SEPARATOR = ";"


def run(book: LiveWorkbook, ref: ReferenceData) -> RunResult:
    result = RunResult(TITLE)
    sheet = book.log_sheet

    if not ref.previous:
        return result.fail(
            f"Table '{config.TABLE_PREV_ACTIVITIES}' not found. "
            "It holds each activity and the activities that must precede it."
        )

    register = book.sheet(config.REGISTER_SHEET)
    if register is None:
        return result.fail(f"Sheet '{config.REGISTER_SHEET}' not found - run Import WIR Data first.")

    register_last = book.last_row(register, config.Register.WIR_NO)
    register_rows = book.block(register, config.REGISTER_FIRST_ROW,
                               register_last, config.Register.LAST_COL)

    approved = _index_approved(register_rows, ref)

    last_row = book.last_row(sheet, config.Main.WIR_NO)
    if last_row < config.FIRST_DATA_ROW:
        return result.fail(f"No rows found in column B of '{sheet.Name}'.")

    rows = book.block(sheet, config.FIRST_DATA_ROW, last_row, config.Main.LAST_COL)

    column: list[str] = []
    filled = empty = no_predecessors = 0
    skipped_hidden = skipped_incomplete = skipped_done = 0

    for offset, values in enumerate(rows):
        row_no = config.FIRST_DATA_ROW + offset

        # Seeded with what is already there. A row this pass skips - hidden,
        # or missing the fields to match on - keeps its existing paths;
        # only a row that was actually looked at gets rewritten.
        column.append(text_of(values[config.Main.ATT_FIRST - 1]))

        if sheet.Rows(row_no).Hidden:
            skipped_hidden += 1
            continue

        if text_of(values[config.Main.STATUS - 1]).lower() == config.STATUS_DONE.lower():
            skipped_done += 1
            continue

        package = text_of(values[config.Main.PACKAGE - 1])
        building = text_of(values[config.Main.BUILDING - 1])
        floor = text_of(values[config.Main.FLOOR - 1])
        activity = text_of(values[config.Main.ACTIVITY - 1])
        if not (package and building and floor and activity):
            skipped_incomplete += 1
            continue

        wanted = set(ref.previous.get(activity, set()))
        if not wanted:
            # Nothing has to come before this activity, so there is
            # nothing to look up and nothing to say. Whatever is in the
            # cell was put there by someone or by an earlier run against
            # a different table - it is not this pass's to delete.
            no_predecessors += 1
            continue

        # A parent counts through its subs, and a sub through its parent.
        for name in list(wanted):
            wanted.update(ref.split_subs.get(name, set()))
            wanted.update(ref.split_parents.get(name, set()))

        tokens = [match_key(t) for t in tokenize(text_of(values[config.Main.AREA - 1]))]
        paths: list[str] = []

        for previous in sorted(wanted):
            key = f"{package}|{building}|{floor}|{previous}".upper()
            for entry_area, path in approved.get(key, []):
                if not path or path in paths:
                    continue
                if _area_accepts(tokens, entry_area):
                    paths.append(path)

        column[-1] = PATH_SEPARATOR.join(paths)
        if paths:
            filled += 1
        else:
            empty += 1

    book.write_column(sheet, config.Main.ATT_FIRST, config.FIRST_DATA_ROW, column)

    result.line(f"Rows given paths:      {filled}")
    result.line(f"Rows with none found:  {empty}")
    result.line(f"No predecessors:       {no_predecessors}")
    result.line(f"Skipped (Done):        {skipped_done}")
    result.line(f"Skipped (hidden):      {skipped_hidden}")
    result.line(f"Skipped (incomplete):  {skipped_incomplete}")
    result.line("")
    result.line("Written to column AA. Rows that were skipped keep what they had.")
    return result


def _index_approved(rows: list[list], ref: ReferenceData) -> dict[str, list[tuple[str, str]]]:
    index: dict[str, list[tuple[str, str]]] = defaultdict(list)
    for row in rows:
        if len(row) < config.Register.LAST_COL:
            row = row + [None] * (config.Register.LAST_COL - len(row))

        if not ref.is_approved(text_of(row[config.Register.STATUS - 1])):
            continue

        package = text_of(row[config.Register.PACKAGE - 1])
        building = text_of(row[config.Register.BUILDING - 1])
        floor = text_of(row[config.Register.FLOOR - 1])
        activity = text_of(row[config.Register.ACTIVITY - 1])
        if not (package and building and floor and activity):
            continue

        key = f"{package}|{building}|{floor}|{activity}".upper()
        index[key].append((
            text_of(row[config.Register.AREA - 1]),
            # The local extract path, not the network one. This is the
            # copy that opens, and it is what the log has always carried.
            text_of(row[config.Register.LOCAL_PATH - 1]),
        ))
    return index


def _area_accepts(tokens: list[str], entry_area: str) -> bool:
    """Blank log area -> any approved WIR counts. Blank register area ->
    that WIR covered everything. "All Apartments" counts when the log row
    names a unit."""
    if not tokens:
        return True
    if not entry_area.strip():
        return True
    if "ALL APARTMENT" in entry_area.upper():
        return any(is_unit_code(t) for t in tokens)

    other = {match_key(t) for t in tokenize(entry_area)}
    return bool(other & set(tokens))
