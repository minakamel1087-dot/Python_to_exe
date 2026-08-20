"""
The PreFlight sheet.

One line per log row, not per finding. A row with a duplicate WIR number
AND a precedence gap is an ERROR row: the worst finding sets the severity,
and the comment carries every finding for that row with the errors first,
so the blocking problem is never buried under a warning.
"""

from __future__ import annotations

from collections import defaultdict
from datetime import datetime

from . import config
from .findings import Severity

HEADERS = [
    "Row", "WIR No.", "Dep.", "Package", "Building", "Floor", "Area",
    "Activity", "Severity", "Check", "What to fix",
]

COL_SEVERITY = 9
COL_CHECK = 10
COL_MESSAGE = 11

_HEADER_FILL = 8021031     # BGR for RGB(0, 70, 127)
_WHITE = 16777215
_ERROR_TEXT = 393372       # RGB(156, 0, 6)
_WARN_TEXT = 6636700       # RGB(156, 101, 0)


def write_report(book, log_sheet_name: str, checked: int, skipped: int, state) -> None:
    sheet = book.ensure_sheet(config.REPORT_SHEET)
    sheet.Cells.Clear()

    errors = sum(1 for f in state.findings if f.severity is Severity.ERROR)
    warnings = len(state.findings) - errors

    title = sheet.Range("A1")
    title.Value = (
        f"Pre-flight: {log_sheet_name}   |   "
        f"{datetime.now():%Y-%m-%d %H:%M}   |   "
        f"checked {checked} row(s), skipped {skipped} marked Done   |   "
        f"{errors} error(s), {warnings} warning(s)"
    )
    title.Font.Bold = True

    header_row = _write_attachment_coverage(sheet, state, checked)
    _write_headers(sheet, header_row)

    if not state.findings:
        cell = sheet.Cells(header_row + 1, 1)
        cell.Value = "No problems found."
        cell.Font.Italic = True
        _finish(sheet, header_row, 0)
        return

    rows = _group_by_row(state)
    block = []
    for row_no, findings in rows:
        context = state.row_context.get(row_no, ("", "", "", "", "", ""))
        severity = state.row_severity[row_no]

        checks: list[str] = []
        chunks: list[str] = []
        for wanted in (Severity.ERROR, Severity.WARN):
            for finding in findings:
                if finding.severity is not wanted:
                    continue
                chunks.append(f"{finding.severity.label} | {finding.check}\n{finding.message}")
                if finding.check not in checks:
                    checks.append(finding.check)

        block.append([
            row_no,
            findings[0].wir_no,
            *context,
            severity.label,
            ", ".join(checks),
            "\n\n".join(chunks),
        ])

    first = header_row + 1
    target = sheet.Range(
        sheet.Cells(first, 1),
        sheet.Cells(first + len(block) - 1, len(HEADERS)),
    )
    target.Value = block

    for offset, (row_no, _) in enumerate(rows):
        cell = sheet.Cells(first + offset, COL_SEVERITY)
        if state.row_severity[row_no] is Severity.ERROR:
            cell.Interior.Color = config.COLOUR_ERROR
            cell.Font.Color = _ERROR_TEXT
        else:
            cell.Interior.Color = config.COLOUR_WARN
            cell.Font.Color = _WARN_TEXT

        anchor = sheet.Cells(first + offset, 1)
        sheet.Hyperlinks.Add(
            Anchor=anchor,
            Address="",
            SubAddress=f"'{log_sheet_name}'!A{row_no}",
            TextToDisplay=str(row_no),
        )

    _apply_filter(sheet, header_row, len(block))

    _finish(sheet, header_row, len(block))


def _apply_filter(sheet, header_row: int, count: int) -> None:
    """Puts the filter arrows on the header row.

    Range.AutoFilter() with no arguments *toggles*. Applied a second time
    over a different range - which happens whenever the number of findings
    changes between runs - Excel raises "AutoFilter method of Range class
    failed" and the whole run dies after the report is already written.

    So any existing filter is cleared first, and the whole thing is
    non-fatal: filter arrows are decoration, and losing them is not a
    reason to throw away a report the user is waiting for.
    """
    if not count:
        return

    try:
        if sheet.AutoFilterMode:
            sheet.AutoFilterMode = False
        sheet.Range(
            sheet.Cells(header_row, 1),
            sheet.Cells(header_row + count, len(HEADERS)),
        ).AutoFilter()
    except Exception:                              # noqa: BLE001
        pass


def _group_by_row(state) -> list[tuple[int, list]]:
    """Error rows first, each block still in log order."""
    grouped: dict[int, list] = defaultdict(list)
    order: list[int] = []
    for finding in state.findings:
        if finding.row not in grouped:
            order.append(finding.row)
        grouped[finding.row].append(finding)

    ranked = [r for r in order if state.row_severity[r] is Severity.ERROR]
    ranked += [r for r in order if state.row_severity[r] is not Severity.ERROR]
    return [(row, grouped[row]) for row in ranked]


def _write_attachment_coverage(sheet, state, checked: int) -> int:
    """A count, not one warning per row. The QC column points at a folder
    that only exists once QC has filed something, so a missing one is
    normal rather than a fault."""
    if not getattr(state, "share_ok", True):
        sheet.Range("A3").Value = (
            "Attachment coverage - SKIPPED: the attachment share is not "
            "reachable, so no path could be checked."
        )
        sheet.Range("A3").Font.Bold = True
        sheet.Range("A3").Font.Color = _WARN_TEXT
        return 5

    sheet.Range("A3").Value = "Attachment coverage (rows where a path was given)"
    sheet.Range("A3").Font.Bold = True

    row = 4
    for label, (found, given) in state.attachments.items():
        sheet.Cells(row, 1).Value = label
        sheet.Cells(row, 3).Value = f"{found} / {checked} found"
        if given == 0:
            sheet.Cells(row, 4).Value = "no paths given on any row"
        elif given > found:
            sheet.Cells(row, 4).Value = f"{given - found} row(s) with nothing to copy"
        row += 1

    return row + 1          # one blank line before the findings


def _write_headers(sheet, header_row: int) -> None:
    target = sheet.Range(
        sheet.Cells(header_row, 1), sheet.Cells(header_row, len(HEADERS))
    )
    target.Value = [HEADERS]
    target.Font.Bold = True
    target.Font.Color = _WHITE
    target.Interior.Color = _HEADER_FILL


def _finish(sheet, header_row: int, count: int) -> None:
    widths = [7, 44, 7, 9, 12, 7, 30, 20, 10, 18, 95]
    for index, width in enumerate(widths, start=1):
        sheet.Columns(index).ColumnWidth = width

    sheet.Columns(7).WrapText = True
    sheet.Columns(COL_MESSAGE).WrapText = True
    sheet.Columns(COL_MESSAGE).VerticalAlignment = -4160    # xlTop

    if count:
        sheet.Range(
            sheet.Cells(header_row + 1, 1),
            sheet.Cells(header_row + count, len(HEADERS)),
        ).EntireRow.AutoFit()
