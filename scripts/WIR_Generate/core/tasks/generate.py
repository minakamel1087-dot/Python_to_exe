"""
Generate WIRs - the whole job, in Python.

For every pending row: render the cover sheet, render the checklist named
in column P, copy in whatever the attachment columns point at, then merge
the lot into one PDF per WIR.

Excel still does the rendering, because ExportAsFixedFormat is how an
Excel sheet becomes a PDF and nothing else reproduces the form layout.
That is COM automation, not a macro - there is no VBA in the workbook.
"""

from __future__ import annotations

import os
import shutil

from .. import config
from ..findings import RunResult
from ..paths import exists, repair, split_paths, trim_slashes, unquote
from ..pdf import merge_folder
from ..workbook import LiveWorkbook, text_of

TITLE = "Generate WIRs"

FORM_SHEET = "WIR-Form"
FORM_INPUT_CELL = "BB1"      # the one cell that drives every field on the form
OUTPUT_FOLDER = "WIRs"
COMMENTS_FORM = "Comments_Form"

XL_TYPE_PDF = 0

# Attachment columns, in the order they should be numbered on disk.
ATTACHMENTS = [
    (config.Main.ATT_QC, "-XTTCH1"),
    (config.Main.ATT_SITE, "-XTTCH2"),
    (config.Main.ATT_PREV, "-XTTCH3"),
    (config.Main.ATT_FIRST, "-XTTCH4"),
]


def run(book: LiveWorkbook, output_root: str = "") -> RunResult:
    result = RunResult(TITLE)
    sheet = book.log_sheet

    form = book.sheet(FORM_SHEET)
    if form is None:
        return result.fail(f"Sheet '{FORM_SHEET}' not found in this workbook.")

    if not output_root:
        output_root = os.path.join(os.path.dirname(book.wb.FullName), OUTPUT_FOLDER)
    os.makedirs(output_root, exist_ok=True)

    last_row = book.last_row(sheet, config.Main.WIR_NO)
    if last_row < config.FIRST_DATA_ROW:
        return result.fail(f"No rows found in column B of '{sheet.Name}'.")

    rows = book.block(sheet, config.FIRST_DATA_ROW, last_row, config.Main.LAST_COL)

    covers = checklists = merged = skipped = failed = 0
    copied = attach_failed = 0
    problems: list[str] = []

    for offset, values in enumerate(rows):
        row_no = config.FIRST_DATA_ROW + offset
        wir_no = text_of(values[config.Main.WIR_NO - 1])
        if not wir_no:
            continue

        if text_of(values[config.Main.STATUS - 1]).lower() == config.STATUS_DONE.lower():
            skipped += 1
            continue

        rev_text = text_of(values[config.Main.REV - 1])
        rev = int(float(rev_text)) if rev_text.replace(".", "").isdigit() else 0
        group = f"{wir_no}-R{rev:02d}"
        row_folder = os.path.join(output_root, group)

        try:
            os.makedirs(row_folder, exist_ok=True)
        except OSError as exc:
            failed += 1
            problems.append(f"Row {row_no}: cannot create folder ({exc})")
            continue

        # -- cover -------------------------------------------------------
        form.Range(FORM_INPUT_CELL).Value = wir_no
        book.excel.Calculate()
        cover = os.path.normpath(os.path.join(row_folder, f"{group}.pdf"))
        try:
            form.ExportAsFixedFormat(XL_TYPE_PDF, cover)
            covers += 1
        except Exception as exc:                   # noqa: BLE001
            failed += 1
            problems.append(f"Row {row_no}: cover PDF failed ({exc})")
            continue

        # -- checklist ---------------------------------------------------
        form_name = text_of(values[config.Main.FORM - 1])
        if form_name:
            checklist = book.sheet(form_name)
            if checklist is None:
                problems.append(f"Row {row_no}: checklist sheet '{form_name}' not found")
            else:
                checklist.Range(FORM_INPUT_CELL).Value = wir_no

                # Comments_Form grows with its content, so its print area is
                # set per run rather than being fixed like the other forms.
                if checklist.Name == COMMENTS_FORM:
                    last = book.last_row(checklist, 1) or 1
                    checklist.PageSetup.PrintArea = f"A1:J{last}"

                book.excel.Calculate()
                target = os.path.normpath(os.path.join(row_folder, f"{wir_no}-XCKL.pdf"))
                try:
                    checklist.ExportAsFixedFormat(XL_TYPE_PDF, target)
                    checklists += 1
                except Exception as exc:           # noqa: BLE001
                    problems.append(f"Row {row_no}: checklist PDF failed ({exc})")

        # -- attachments -------------------------------------------------
        for column, suffix in ATTACHMENTS:
            ok, bad = _copy_attachments(
                text_of(values[column - 1]), row_folder, wir_no, suffix
            )
            copied += ok
            attach_failed += bad

        # -- one PDF per WIR ---------------------------------------------
        outcome = merge_folder(row_folder, output_root, group)
        if outcome.ok:
            merged += 1
            if outcome.kept_files:
                problems.append(
                    f"Row {row_no}: merged {outcome.merged_count} PDF(s); "
                    f"kept {len(outcome.kept_files)} non-PDF file(s) in {group}"
                )
        else:
            problems.append(f"Row {row_no}: merge failed ({outcome.error})")

        sheet.Cells(row_no, config.Main.STATUS).Value = config.STATUS_DONE

    result.line(f"Covers rendered:   {covers}")
    result.line(f"Checklists:        {checklists}")
    result.line(f"Merged into one:   {merged}")
    result.line(f"Attachments:       {copied} copied, {attach_failed} failed")
    result.line(f"Skipped (Done):    {skipped}")
    result.line(f"Failed rows:       {failed}")
    result.line("")
    result.line(f"Output: {output_root}")

    if problems:
        result.line("")
        result.line("Problems:")
        result.details.extend(f"  {p}" for p in problems[:20])
        if len(problems) > 20:
            result.line(f"  ... and {len(problems) - 20} more")

    return result


def _copy_attachments(raw: str, row_folder: str, wir_no: str,
                      suffix: str) -> tuple[int, int]:
    """Copies whatever a path cell points at. The cell may hold several
    paths, and each may be a file or a folder - a folder contributes all
    of its files, numbering continuing across the cell."""
    if not raw or raw.startswith("#"):
        return 0, 0

    copied = failed = 0
    index = 0

    for original in split_paths(raw):
        source = trim_slashes(unquote(original))
        if not exists(source):
            repaired = repair(original)
            if exists(repaired):
                source = repaired
            else:
                failed += 1
                continue

        if os.path.isfile(source):
            index += 1
            if _copy_one(source, row_folder, wir_no, suffix, index):
                copied += 1
            else:
                failed += 1
            continue

        try:
            names = sorted(os.listdir(source))
        except OSError:
            failed += 1
            continue

        files = [n for n in names if os.path.isfile(os.path.join(source, n))]
        if not files:
            failed += 1
            continue

        for name in files:
            index += 1
            if _copy_one(os.path.join(source, name), row_folder, wir_no, suffix, index):
                copied += 1
            else:
                failed += 1

    return copied, failed


def _copy_one(source: str, row_folder: str, wir_no: str, suffix: str, index: int) -> bool:
    extension = os.path.splitext(source)[1]
    target = os.path.join(row_folder, f"{wir_no}{suffix}-{index}{extension}")
    try:
        shutil.copyfile(source, target)
        return True
    except OSError:
        return False
