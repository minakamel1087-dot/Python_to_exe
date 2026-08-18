"""
Import WIR Data - reloads the WIRs sheet from the master register.

The register is a separate workbook with one sheet per trade. Rows A:T of
each are stacked into the WIRs sheet from row 4, and the four derived
columns (AA..AD) are computed here rather than left as formulas - the
workbook has no macros now, and AD used to depend on one.
"""

from __future__ import annotations

import os

from .. import config
from ..findings import RunResult
from ..workbook import LiveWorkbook, find_open_workbook, text_of
from . import fix_prefix

TITLE = "Import WIR Data"

SOURCE_SHEETS = ["ELE", "ELV"]
SOURCE_FIRST_ROW = 4
SOURCE_LAST_COL = 20            # A:T

HEADER_ROW = 3
FIRST_TARGET_ROW = 4
CLEAR_COLUMNS = 30              # A:AD

# The local folder the commented WIRs are extracted to.
LOCAL_PATH = config.REGISTER_LOCAL_PREFIX


def run(book: LiveWorkbook, source_path: str,
        source_action: str = "discard") -> RunResult:
    """source_action decides what happens to the register file afterwards,
    and only matters when the user already had it open: "save" closes it
    keeping their changes, "discard" closes it throwing them away. When we
    opened it ourselves it is read-only, so there is nothing to save."""
    result = RunResult(TITLE)

    if not source_path:
        return result.fail("No register file chosen.")
    if not os.path.isfile(source_path):
        return result.fail(f"Register file not found:\n{source_path}")

    target = book.sheet(config.REGISTER_SHEET)
    if target is None:
        return result.fail(f"Sheet '{config.REGISTER_SHEET}' not found in this workbook.")

    # Was it already open? Workbooks.Open would hand back that same
    # workbook, and closing it afterwards would take the user's unsaved
    # work with it.
    already_open = find_open_workbook(book.excel, source_path) is not None
    source = book.excel.Workbooks.Open(source_path, ReadOnly=True)
    try:
        target.Range(
            target.Cells(1, 1),
            target.Cells(target.Rows.Count, CLEAR_COLUMNS),
        ).ClearContents()

        target.Cells(HEADER_ROW, 27).Value = "WIR No and Status"
        target.Cells(HEADER_ROW, 28).Value = "Code"
        target.Cells(HEADER_ROW, 29).Value = "Area"
        target.Cells(HEADER_ROW, 30).Value = "Local Path"

        next_row = FIRST_TARGET_ROW
        imported: list[str] = []

        for name in SOURCE_SHEETS:
            sheet = _sheet(source, name)
            if sheet is None:
                imported.append(f"{name}: sheet not found, skipped")
                continue

            if sheet.FilterMode:
                sheet.ShowAllData()
            if sheet.AutoFilterMode:
                sheet.AutoFilterMode = False

            last = sheet.Cells(sheet.Rows.Count, 1).End(-4162).Row
            if last < SOURCE_FIRST_ROW:
                imported.append(f"{name}: nothing to copy")
                continue

            block = sheet.Range(
                sheet.Cells(SOURCE_FIRST_ROW, 1),
                sheet.Cells(last, SOURCE_LAST_COL),
            ).Value2
            rows = [list(r) for r in block] if isinstance(block, tuple) else [[block]]

            target.Range(
                target.Cells(next_row, 1),
                target.Cells(next_row + len(rows) - 1, SOURCE_LAST_COL),
            ).Value = rows

            derived = [_derive(row) for row in rows]
            target.Range(
                target.Cells(next_row, 27),
                target.Cells(next_row + len(derived) - 1, 30),
            ).Value = derived

            next_row += len(rows)
            imported.append(f"{name}: {len(rows)} row(s)")

    finally:
        # Only a workbook the user already had open can have changes worth
        # keeping; ours was opened read-only.
        save = already_open and source_action == "save"
        source.Close(SaveChanges=save)

    # The paths just imported all point at the share; the local extract
    # folder is what opens. Nothing else puts them there, so this runs
    # here rather than waiting to be remembered.
    fix_prefix.apply(book, result)

    result.line("")
    result.line(f"Register: {os.path.basename(source_path)}")
    for line in imported:
        result.line(f"  {line}")
    result.line("")
    result.line(f"WIRs sheet filled to row {next_row - 1}.")
    return result


def _sheet(workbook, name: str):
    for sheet in workbook.Worksheets:
        if sheet.Name == name:
            return sheet
    return None


def _derive(row: list) -> list:
    """AA = number and status, AB = location code, AC = area,
    AD = the local path the commented PDF actually opens from."""
    wir_no = text_of(row[config.Register.WIR_NO - 1])
    status = text_of(row[config.Register.STATUS - 1])
    package = text_of(row[config.Register.PACKAGE - 1])
    building = text_of(row[config.Register.BUILDING - 1])
    floor = text_of(row[config.Register.FLOOR - 1])
    activity = text_of(row[config.Register.ACTIVITY - 1])
    area = text_of(row[config.Register.AREA - 1])
    path = text_of(row[config.Register.PATH - 1])

    number = wir_no[-6:] if len(wir_no) >= 6 else wir_no
    tail = path[-52:] if len(path) > 52 else path

    code = "-".join(p for p in (package, building, floor, activity) if p)
    return [
        f"{number}({status})" if number else "",
        code,
        area,
        f"{LOCAL_PATH}{tail}" if tail else "",
    ]
