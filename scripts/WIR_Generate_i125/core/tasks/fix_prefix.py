"""
Fix WIR Path Prefix - repoints the register's commented-WIR paths at the
local extract folder.

The register stores each commented WIR against the network share; the
local folder is what actually opens. This is pure text substitution - no
disk access, nothing to find - so it runs automatically at the end of
Import WIR Data, which is the only thing that puts those paths there.
It stays on its own button for the times the register is edited by hand.
"""

from __future__ import annotations

from .. import config
from ..findings import RunResult
from ..workbook import LiveWorkbook, text_of

TITLE = "Fix WIR Path Prefix"


def run(book: LiveWorkbook) -> RunResult:
    result = RunResult(TITLE)
    changed = apply(book, result)
    if changed < 0:
        return result.fail(f"Sheet '{config.REGISTER_SHEET}' not found.")
    return result


def apply(book: LiveWorkbook, result: RunResult) -> int:
    """Does the work and appends its lines to whatever result it is given,
    so Import WIR Data can fold this into its own report.

    Returns the number of paths changed, or -1 if the register is missing.
    """
    sheet = book.sheet(config.REGISTER_SHEET)
    if sheet is None:
        return -1

    last_row = book.last_row(sheet, config.Register.PATH)
    if last_row < config.REGISTER_FIRST_ROW:
        result.line("Register has no paths to repoint.")
        return 0

    rng = sheet.Range(
        sheet.Cells(config.REGISTER_FIRST_ROW, config.Register.PATH),
        sheet.Cells(last_row, config.Register.PATH),
    )
    values = rng.Value2
    rows = [list(r) for r in values] if isinstance(values, tuple) else [[values]]

    network = config.REGISTER_NETWORK_PREFIX.lower()
    changed = 0
    out: list[str] = []

    for row in rows:
        current = text_of(row[0])
        if current and current.lower().startswith(network):
            current = (
                config.REGISTER_LOCAL_PREFIX
                + current[len(config.REGISTER_NETWORK_PREFIX):]
            )
            changed += 1
        out.append(current)

    if changed:
        rng.Value = [[v] for v in out]

    result.line(f"Register paths repointed to local: {changed}")
    return changed
