"""
Fix Links - the old "Fix Attachment Links" and "Fix WIR Path Prefix"
merged, because both do the same job on different columns:

  1. the log's own attachment paths (Main column N)
  2. the register's commented-WIR paths (WIRs column N), which are stored
     against the network share but opened from the local extract folder

Rules, deliberately conservative:
  - A link that already resolves is never touched.
  - A local path (C: D: E: F: G: H:) is never touched.
  - A corrected path is written back ONLY if the files are actually there
    afterwards. If the correction finds nothing either, the original text
    is left exactly as it was and the cell is flagged, because at that
    point it needs a person, not a guess.
  - Multiple paths in one cell keep their original separators.
"""

from __future__ import annotations

from .. import config
from ..findings import RunResult
from ..paths import repair_cell
from ..workbook import LiveWorkbook, text_of

TITLE = "Fix Links"


def run(book: LiveWorkbook) -> RunResult:
    result = RunResult(TITLE)
    sheet = book.log_sheet

    last_row = book.last_row(sheet, config.Main.WIR_NO)
    if last_row < config.FIRST_DATA_ROW:
        return result.fail(f"No rows found in column B of '{sheet.Name}'.")

    _fix_log_links(book, sheet, last_row, result)
    _fix_register_prefix(book, result)
    return result


def _fix_log_links(book: LiveWorkbook, sheet, last_row: int, result: RunResult) -> None:
    rows = book.block(sheet, config.FIRST_DATA_ROW, last_row, config.Main.LAST_COL)

    # Last run's flags go first, or a link fixed since then stays red.
    book.clear_fill(sheet, config.Main.LINK, config.FIRST_DATA_ROW, last_row)

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

    result.line(f"Log rows with a link:  {scanned}")
    result.line(f"Links corrected:       {repaired}")
    result.line(f"Still not found:       {broken}")
    result.line(f"Skipped (Done):        {skipped_done}")
    if problems:
        result.line("")
        result.line("Left as typed and flagged in column N:")
        result.details.extend(f"  {p}" for p in problems)


def _fix_register_prefix(book: LiveWorkbook, result: RunResult) -> None:
    """The register stores commented WIRs against the network share; the
    local extract folder is what actually opens."""
    sheet = book.sheet(config.REGISTER_SHEET)
    if sheet is None:
        result.line("")
        result.line(f"Register sheet '{config.REGISTER_SHEET}' not found - prefix step skipped.")
        return

    last_row = book.last_row(sheet, config.Register.PATH)
    if last_row < config.REGISTER_FIRST_ROW:
        return

    rng = sheet.Range(
        sheet.Cells(config.REGISTER_FIRST_ROW, config.Register.PATH),
        sheet.Cells(last_row, config.Register.PATH),
    )
    values = rng.Value
    rows = [list(r) for r in values] if isinstance(values, tuple) else [[values]]

    changed = 0
    out: list[str] = []
    network = config.REGISTER_NETWORK_PREFIX.lower()

    for row in rows:
        current = text_of(row[0])
        if current and current.lower().startswith(network):
            current = config.REGISTER_LOCAL_PREFIX + current[len(config.REGISTER_NETWORK_PREFIX):]
            changed += 1
        out.append(current)

    if changed:
        rng.Value = [[v] for v in out]

    result.line("")
    result.line(f"Register paths repointed to local: {changed}")
