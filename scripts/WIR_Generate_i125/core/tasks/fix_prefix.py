"""
Fix WIR Path Prefix - fills the register's local column from its server
column.

Column N holds each commented WIR as the register files it, on the share.
Column AD holds the same file under the local extract folder, which is
the copy that actually opens. This reads N, substitutes the server prefix
for its local twin from the **Paths** sheet, and writes the result to AD.

Two things it does not do, both deliberate:

  * **Column N is never written.** An earlier version substituted in
    place, which would have destroyed the register's server paths the
    first time it matched anything. It never did match - the prefix it
    looked for was wrong - so the damage was only ever theoretical.
  * The radio buttons make no difference here. AD is the local column by
    definition; what Server/Local decides is which column other tasks
    *read*, not what this one writes.

Pure text substitution, no disk access, so it runs at the end of Import
WIR Log as well as on its own button.
"""

from __future__ import annotations

from .. import config, pathmap
from ..findings import RunResult
from ..workbook import LiveWorkbook, text_of

TITLE = "Fix WIR Path Prefix"


def run(book: LiveWorkbook) -> RunResult:
    result = RunResult(TITLE)
    changed = apply(book, result)
    if changed < 0:
        return result.fail(f"Sheet '{config.REGISTER_SHEET}' not found.")
    return result


def apply(book: LiveWorkbook, result: RunResult,
          table: pathmap.PathMap | None = None) -> int:
    """Does the work and appends its lines to whatever result it is given,
    so Import WIR Log can fold this into its own report.

    Returns the number of paths written, or -1 if the register is missing.
    """
    sheet = book.sheet(config.REGISTER_SHEET)
    if sheet is None:
        return -1

    table = table if table is not None else pathmap.load(book)
    if not table.pairs:
        result.line(f"No '{config.PATHS_SHEET}' sheet - local paths not written.")
        return 0

    last_row = book.last_row(sheet, config.Register.PATH)
    if last_row < config.REGISTER_FIRST_ROW:
        result.line("Register has no paths to repoint.")
        return 0

    source = book.block_column(sheet, config.Register.PATH,
                               config.REGISTER_FIRST_ROW, last_row)

    written = 0
    warnings: set[str] = set()
    out: list[str] = []

    for value in source:
        server_path = text_of(value)
        if not server_path:
            out.append("")
            continue

        local_path, warning = table.to_local(server_path)
        if warning:
            warnings.add(warning)
        if local_path != server_path:
            written += 1
        out.append(local_path)

    book.write_column(sheet, config.Register.LOCAL_PATH,
                      config.REGISTER_FIRST_ROW, out)

    result.line(f"Local paths written to column AD: {written}")
    for warning in sorted(warnings):
        result.line(f"  {warning}")
    return written
