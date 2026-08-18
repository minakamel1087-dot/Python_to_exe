"""
The tasks the buttons run.

Every one of them works against the live workbook and returns a RunResult
the UI can render. Nothing here calls a VBA macro - the workbook is a
workbook again, and this program is the tool.

Excel still renders the PDFs, because ExportAsFixedFormat is how a sheet
becomes a PDF. That is COM automation from outside, not code living in
the file.
"""

from __future__ import annotations

from ..findings import RunResult
from ..progress import SILENT, Progress
from ..reference import ReferenceData, load
from ..workbook import LiveWorkbook
from . import (
    check_areas,
    clear_sheets,
    extract_covers,
    extract_pdfs,
    fix_attachments,
    fix_prefix,
    generate,
    history,
    import_previous,
    import_register,
    preflight,
)

__all__ = [
    "check_areas", "clear_sheets", "extract_covers", "extract_pdfs",
    "fix_attachments", "fix_prefix", "generate", "generate_wirs", "history",
    "import_previous", "import_register", "preflight", "run_all",
]

# Standalone: no workbook, no sheet, no Excel. The window runs these
# without attaching, so they work with the workbook closed.
NO_WORKBOOK = {"extract_covers"}


def generate_wirs(book: LiveWorkbook, ref: ReferenceData, output: str = "",
                  progress: Progress = SILENT) -> list[RunResult]:
    """Import Previous WIR Paths, then Generate.

    Column AA has to be current before a WIR is built. Its paths are only
    right as of the last register import, and a WIR that goes out without
    its predecessor covers comes back rejected - so Generate refreshes
    them itself rather than relying on anyone remembering.

    The separate button stays, for filling the column without generating.

    A failure in the first step stops the second: generating against
    paths known to be stale is worse than not generating at all.
    """
    progress.report(0, 0, "Refreshing previous activity paths...")
    first = import_previous.run(book, ref)
    if not first.ok:
        return [first]

    return [first, generate.run(book, output, progress=progress)]


def run_all(book: LiveWorkbook, ref: ReferenceData) -> list[RunResult]:
    """The three checks that belong before generating, in order, because
    each feeds the next: repaired links stop the attachment check
    reporting phantom failures, and the area pass puts its suggestions in
    K before pre-flight reads the rows.

    Fix WIR Path Prefix is not here - it belongs to the register, and
    Import WIR Data already runs it.
    """
    results = [fix_attachments.run(book)]
    if not results[0].ok:
        return results

    results.append(check_areas.run(book, ref))

    # The area pass may have written to the sheet; reload the reference
    # tables so pre-flight sees the workbook the user now has.
    results.append(preflight.run(book, load(book)))
    return results
