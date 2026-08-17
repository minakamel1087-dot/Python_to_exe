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
from ..reference import ReferenceData, load
from ..workbook import LiveWorkbook
from . import (
    check_areas,
    clear_sheets,
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
    "check_areas", "clear_sheets", "extract_pdfs", "fix_attachments",
    "fix_prefix", "generate", "history", "import_previous", "import_register",
    "preflight", "run_all",
]


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
