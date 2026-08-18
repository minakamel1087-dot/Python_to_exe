"""
Extract PDFs - copies whatever the visible rows point at into one folder.

The VBA asked you to select a range by hand. This takes the same column
the rest of the tools use and works on the rows you can see, which is the
same filtered set Copy To History uses.
"""

from __future__ import annotations

import os
import shutil

from .. import config
from ..findings import RunResult
from ..paths import exists, repair, split_paths, trim_slashes, unquote
from ..workbook import LiveWorkbook, text_of

TITLE = "Extract PDFs"


def run(book: LiveWorkbook, destination: str, column: int = config.Main.LINK) -> RunResult:
    result = RunResult(TITLE)
    sheet = book.log_sheet

    if not destination:
        return result.fail("No destination folder chosen.")
    os.makedirs(destination, exist_ok=True)

    last_row = book.last_row(sheet, config.Main.WIR_NO)
    if last_row < config.FIRST_DATA_ROW:
        return result.fail(f"No rows found in column B of '{sheet.Name}'.")

    rows = book.block(sheet, config.FIRST_DATA_ROW, last_row, config.Main.LAST_COL)

    copied = failed = skipped = 0
    problems: list[str] = []

    for offset, values in enumerate(rows):
        row_no = config.FIRST_DATA_ROW + offset
        if not text_of(values[config.Main.WIR_NO - 1]):
            continue
        if sheet.Rows(row_no).Hidden:
            skipped += 1
            continue

        raw = text_of(values[column - 1])
        if not raw or raw.startswith("#"):
            continue

        for original in split_paths(raw):
            source = trim_slashes(unquote(original))
            if not exists(source):
                repaired = repair(original)
                if exists(repaired):
                    source = repaired
                else:
                    failed += 1
                    problems.append(f"Row {row_no}: not found - {source}")
                    continue

            if os.path.isfile(source):
                ok = _copy(source, destination)
                copied += ok
                failed += 0 if ok else 1
                continue

            try:
                names = sorted(os.listdir(source))
            except OSError as exc:
                failed += 1
                problems.append(f"Row {row_no}: {exc}")
                continue

            for name in names:
                path = os.path.join(source, name)
                if os.path.isfile(path):
                    ok = _copy(path, destination)
                    copied += ok
                    failed += 0 if ok else 1

    result.line(f"Files copied:      {copied}")
    result.line(f"Failed:            {failed}")
    result.line(f"Rows hidden/skipped: {skipped}")
    result.line("")
    result.line(f"Destination: {destination}")

    if problems:
        result.line("")
        result.details.extend(f"  {p}" for p in problems[:15])

    return result


def _copy(source: str, destination: str) -> int:
    """Keeps the original filename, numbering rather than overwriting when
    the name is already taken."""
    name = os.path.basename(source)
    stem, extension = os.path.splitext(name)
    target = os.path.join(destination, name)

    counter = 1
    while os.path.exists(target):
        target = os.path.join(destination, f"{stem} ({counter}){extension}")
        counter += 1

    try:
        shutil.copyfile(source, target)
        return 1
    except OSError:
        return 0
