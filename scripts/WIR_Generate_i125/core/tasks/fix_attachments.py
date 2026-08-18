"""
Fix Attachment Links - repairs the log's own attachment paths before
anything is generated.

Site attachment paths arrive from other people's machines and are wrong
in three predictable ways:

    "Z:\\WIR request\\..."   wrapped in quotes (Explorer's "Copy as path")
    S:\\WIR request\\...     the share mapped to a different drive letter
    Z:\\WIR request\\...     the \\Common\\ level missing

Rules, deliberately conservative:
  - A link that already resolves is never touched.
  - A local path (C: D: E: F: G: H:) is never touched.
  - A corrected path is written back ONLY if the files are actually
    there afterwards. If the correction finds nothing either, the
    original text is left exactly as it was and the cell is flagged,
    because at that point it needs a person, not a guess.
  - Multiple paths in one cell keep their original separators.

Column N only. U, W, Y and AA are formulas that derive from it.
"""

from __future__ import annotations

from .. import config
from ..findings import RunResult
from ..paths import repair_cell, share_available
from ..workbook import LiveWorkbook, text_of

TITLE = "Fix Attachment Links"


def run(book: LiveWorkbook) -> RunResult:
    result = RunResult(TITLE)
    sheet = book.log_sheet

    last_row = book.last_row(sheet, config.Main.WIR_NO)
    if last_row < config.FIRST_DATA_ROW:
        return result.fail(f"No rows found in column B of '{sheet.Name}'.")

    rows = book.block(sheet, config.FIRST_DATA_ROW, last_row, config.Main.LAST_COL)

    # Last run's flags go first, or a link fixed since then stays red.
    book.clear_fill(sheet, config.Main.LINK, config.FIRST_DATA_ROW, last_row)

    if not share_available():
        # Without the share every path under it looks broken. Flagging the
        # whole column then is noise, and it hides the rows that are
        # genuinely wrong once the connection is back.
        result.line(f"Share {config.SHARE_ROOT} is not reachable - link check skipped.")
        result.line("Connect to the server and run this again.")
        return result

    scanned = repaired = broken = skipped_done = 0
    problems: list[str] = []
    updates: dict[int, str] = {}

    for offset, values in enumerate(rows):
        row_no = config.FIRST_DATA_ROW + offset
        wir_no = text_of(values[config.Main.WIR_NO - 1])
        if not wir_no:
            continue

        if text_of(values[config.Main.STATUS - 1]).lower() == config.STATUS_DONE.lower():
            skipped_done += 1
            continue

        raw = text_of(values[config.Main.LINK - 1])
        if not raw:
            continue

        scanned += 1
        new_text, changed, still_broken = repair_cell(raw)

        if changed:
            updates[row_no] = new_text
            repaired += 1
        if still_broken:
            broken += 1
            book.fill(sheet, row_no, config.Main.LINK, config.COLOUR_ERROR)
            if len(problems) < 15:
                problems.append(f"Row {row_no}: {still_broken[0]}")

    for row_no, text in updates.items():
        sheet.Cells(row_no, config.Main.LINK).Value = text

    result.line(f"Rows with a link:  {scanned}")
    result.line(f"Links corrected:   {repaired}")
    result.line(f"Still not found:   {broken}")
    result.line(f"Skipped (Done):    {skipped_done}")

    if problems:
        result.line("")
        result.line("Left as typed and flagged in column N:")
        result.details.extend(f"  {p}" for p in problems)

    return result
