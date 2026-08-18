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
from ..progress import SILENT, Progress
from ..workbook import LiveWorkbook, text_of

TITLE = "Generate WIRs"

FORM_SHEET = "WIR-Form"
FORM_INPUT_CELL = "BB1"      # the one cell that drives every field on the form
OUTPUT_FOLDER = "WIRs"
COMMENTS_FORM = "Comments_Form"

XL_TYPE_PDF = 0

# Attachment columns, in the order they should be numbered on disk, with
# how each behaves:
#
#   first_page  take only the cover. ATT_FIRST is column AA, the approved
#               predecessor WIRs - whole commented documents of which only
#               the first page is wanted. (ATT_PREV, despite the name, is
#               column Y: the earlier revision of this same WIR, kept
#               whole.) Output naming is unchanged either way.
#
#   required    a path that does not resolve is a failure. QC is the
#               exception: those files are often filed after the WIR is
#               raised, so a missing one is reported and counted, not
#               treated as an error.
ATTACHMENTS = [
    (config.Main.ATT_QC, "-XTTCH1", False, False),
    (config.Main.ATT_SITE, "-XTTCH2", False, True),
    (config.Main.ATT_PREV, "-XTTCH3", False, True),
    (config.Main.ATT_FIRST, "-XTTCH4", True, True),
]


def run(book: LiveWorkbook, output_root: str = "",
        progress: Progress = SILENT) -> RunResult:
    result = RunResult(TITLE)
    sheet = book.log_sheet

    form = book.sheet(FORM_SHEET)
    if form is None:
        return result.fail(f"Sheet '{FORM_SHEET}' not found in this workbook.")

    # The WIRs folder beside the workbook. Not asked for - it is always
    # the same place, and being asked every time is just a step to get
    # wrong. The argument exists so tests can write somewhere else.
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
    qc_missing: list[str] = []      # reported, never counted as failures
    cancelled = False

    # Counted up front so the bar means something. Done and blank rows are
    # not work, and including them would make it crawl then leap.
    pending = sum(
        1 for v in rows
        if text_of(v[config.Main.WIR_NO - 1])
        and text_of(v[config.Main.STATUS - 1]).lower() != config.STATUS_DONE.lower()
    )
    done_so_far = 0

    for offset, values in enumerate(rows):
        row_no = config.FIRST_DATA_ROW + offset
        wir_no = text_of(values[config.Main.WIR_NO - 1])
        if not wir_no:
            continue

        if text_of(values[config.Main.STATUS - 1]).lower() == config.STATUS_DONE.lower():
            skipped += 1
            continue

        # Between rows, never inside one: a WIR is finished or it is not
        # started, and a half-built folder marked Done would be submitted.
        if progress.cancelled:
            cancelled = True
            break

        done_so_far += 1
        progress.report(done_so_far, pending, wir_no)

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
        for column, suffix, first_page_only, required in ATTACHMENTS:
            ok, bad, issues, notes = _copy_attachments(
                text_of(values[column - 1]), row_folder, wir_no, suffix,
                first_page_only, required,
            )
            copied += ok
            attach_failed += bad
            problems.extend(f"Row {row_no}: {issue}" for issue in issues)
            for note in notes:
                qc_missing.append(f"Row {row_no}  {wir_no}")
                qc_missing.append(f"      {note}")

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

    if cancelled:
        result.line("STOPPED by you. Rows already generated are marked Done "
                    "and keep their PDFs; the rest were not started.")
        result.line("")

    result.line(f"Covers rendered:   {covers}")
    result.line(f"Checklists:        {checklists}")
    result.line(f"Merged into one:   {merged}")
    result.line(f"Attachments:       {copied} copied, {attach_failed} failed")
    if qc_missing:
        # Two lines are appended per row, so halve to count rows.
        result.line(f"QC attachment:     {len(qc_missing) // 2} of {done_so_far} "
                    f"row(s) had none - not an error")
    result.line(f"Skipped (Done):    {skipped}")
    result.line(f"Failed rows:       {failed}")
    result.line("")
    result.line(f"Output: {output_root}")

    if qc_missing:
        result.line("")
        result.line("QC attachment not found (WIR generated without it):")
        for line in qc_missing:
            result.line(f"  {line}")

    if problems:
        result.line("")
        result.line("Problems:")
        result.details.extend(f"  {p}" for p in problems)

    progress.report(done_so_far, pending, "Stopped" if cancelled else "Done")
    return result


# Windows drops these into shared folders. They are not attachments, and
# copying them also stops the row folder being cleaned up after the merge
# because something non-PDF is left in it.
JUNK_FILES = {"thumbs.db", "desktop.ini", ".ds_store", "ehthumbs.db"}

FILE_ATTRIBUTE_HIDDEN = 0x02
FILE_ATTRIBUTE_SYSTEM = 0x04


def _is_junk(path: str) -> bool:
    if os.path.basename(path).lower() in JUNK_FILES:
        return True
    try:
        attrs = os.stat(path).st_file_attributes      # Windows only
    except (OSError, AttributeError):
        return False
    return bool(attrs & (FILE_ATTRIBUTE_HIDDEN | FILE_ATTRIBUTE_SYSTEM))


def _copy_attachments(raw: str, row_folder: str, wir_no: str, suffix: str,
                      first_page_only: bool = False,
                      required: bool = True
                      ) -> tuple[int, int, list[str], list[str]]:
    """Copies whatever a path cell points at. The cell may hold several
    paths, and each may be a file or a folder - a folder contributes all
    of its files, numbering continuing across the cell.

    Returns copied, failed, problems, notes. When `required` is false a
    path that does not resolve becomes a note instead of a failure - the
    WIR is still perfectly generatable without it."""
    if not raw or raw.startswith("#"):
        return 0, 0, [], []

    copied = failed = 0
    index = 0
    problems: list[str] = []
    notes: list[str] = []

    for original in split_paths(raw):
        source = trim_slashes(unquote(original))
        if not exists(source):
            repaired = repair(original)
            if exists(repaired):
                source = repaired
            elif required:
                failed += 1
                problems.append(f"not found: {source}")
                continue
            else:
                notes.append(source)
                continue

        if os.path.isfile(source):
            index += 1
            if _copy_one(source, row_folder, wir_no, suffix, index, first_page_only):
                copied += 1
            else:
                failed += 1
                problems.append(f"could not copy: {source}")
            continue

        try:
            names = sorted(os.listdir(source))
        except OSError as exc:
            failed += 1
            problems.append(f"{source}: {exc}")
            continue

        files = [
            n for n in names
            if os.path.isfile(os.path.join(source, n))
            and not _is_junk(os.path.join(source, n))
        ]
        if not files:
            failed += 1
            problems.append(f"nothing to copy in: {source}")
            continue

        for name in files:
            index += 1
            if _copy_one(os.path.join(source, name), row_folder, wir_no, suffix,
                         index, first_page_only):
                copied += 1
            else:
                failed += 1
                problems.append(f"could not copy: {os.path.join(source, name)}")

    return copied, failed, problems, notes


def _copy_one(source: str, row_folder: str, wir_no: str, suffix: str, index: int,
              first_page_only: bool = False) -> bool:
    extension = os.path.splitext(source)[1]
    target = os.path.join(row_folder, f"{wir_no}{suffix}-{index}{extension}")

    if first_page_only and extension.lower() == ".pdf":
        return _copy_first_page(source, target)

    try:
        shutil.copyfile(source, target)
        return True
    except OSError:
        return False


def _copy_first_page(source: str, target: str) -> bool:
    """The cover of a commented WIR, not the whole thing.

    Written under a temporary name and renamed only once complete - a
    part-written PDF looks like a finished one and would be merged into
    the WIR as if it were fine.

    Falls back to copying the file whole if it cannot be read: a WIR with
    too many pages in it is a nuisance, a WIR missing its evidence is a
    rejection.
    """
    partial = target + ".part"
    try:
        from pypdf import PdfReader, PdfWriter

        reader = PdfReader(source)
        if not reader.pages:
            raise ValueError("no pages")

        writer = PdfWriter()
        writer.add_page(reader.pages[0])
        with open(partial, "wb") as handle:
            writer.write(handle)
        writer.close()
        os.replace(partial, target)
        return True
    except Exception:                              # noqa: BLE001
        try:
            if os.path.exists(partial):
                os.remove(partial)
        except OSError:
            pass
        try:
            shutil.copyfile(source, target)
            return True
        except OSError:
            return False
