"""
Extract WIR Cover Page.

Copies the **first page only** of every commented WIR PDF from the share
into the local extract folder. A file already in the destination is left
alone, so a second run only picks up what is new.

Deliberately standalone: it touches no workbook and no sheet, needs no
Excel, and shares nothing with the log tasks but the result type. The
destination is the folder the register's links point at, which is why it
lives beside them rather than in a tool of its own.
"""

from __future__ import annotations

import os

from .. import config
from ..findings import RunResult
from ..paths import folder_reachable
from ..progress import SILENT, Progress

TITLE = "Extract WIR Cover Page"

# A part-written PDF looks like a finished one. Each file is built under
# this suffix and only renamed once it is complete, so an interrupted run
# leaves nothing that a later run would mistake for done.
PARTIAL = ".part"


def run(source: str = "", destination: str = "",
        progress: Progress = SILENT) -> RunResult:
    result = RunResult(TITLE)

    source = os.path.normpath(source or config.COVER_SOURCE)
    destination = os.path.normpath(destination or config.COVER_DEST)

    progress.report(0, 0, "Looking for the share...")
    if not folder_reachable(source):
        return result.fail(
            f"Source folder is not reachable:\n{source}\n\n"
            "Connect to the server and try again."
        )

    try:
        os.makedirs(destination, exist_ok=True)
    except OSError as exc:
        return result.fail(f"Cannot create the destination folder:\n{destination}\n\n{exc}")

    try:
        names = sorted(os.listdir(source))
    except OSError as exc:
        return result.fail(f"Cannot read the source folder:\n{source}\n\n{exc}")

    from pypdf import PdfReader, PdfWriter          # imported late; only this task needs it

    pdfs = [n for n in names if n.lower().endswith(".pdf")]
    extracted = skipped = processed = 0
    failures: list[str] = []
    cancelled = False

    for index, name in enumerate(pdfs, start=1):
        if progress.cancelled:
            cancelled = True
            break
        processed = index

        progress.report(index, len(pdfs), name)

        target = os.path.join(destination, name)
        if os.path.exists(target):
            skipped += 1
            continue

        partial = target + PARTIAL
        try:
            reader = PdfReader(os.path.join(source, name))
            if not reader.pages:
                failures.append(f"{name}: no pages")
                continue

            writer = PdfWriter()
            writer.add_page(reader.pages[0])
            with open(partial, "wb") as handle:
                writer.write(handle)
            writer.close()

            os.replace(partial, target)
            extracted += 1
        except Exception as exc:                    # noqa: BLE001 - reported, not raised
            failures.append(f"{name}: {exc}")
            try:
                if os.path.exists(partial):
                    os.remove(partial)
            except OSError:
                pass

    result.line(f"From: {source}")
    result.line(f"To:   {destination}")
    result.line("")
    if cancelled:
        result.line("STOPPED by you - what had been written is kept.")
        result.line("")
    result.line(f"PDFs on the share:   {len(pdfs)}")
    result.line(f"Covers extracted:    {extracted}")
    result.line(f"Already there:       {skipped}")

    other = len(names) - len(pdfs)
    if other:
        result.line(f"Not a PDF, ignored:  {other}")

    if failures:
        # Every one of them, named. A count alone cannot be acted on, and
        # these are the files someone has to go and look at.
        result.line("")
        result.line(f"Could not be read:   {len(failures)}")
        result.line("")
        for line in failures:
            result.line(f"  {line}")
        result.line("")
        result.line("These were left on the share untouched.")

    # Where it actually got to. Reporting the full count after a stop
    # would fill the bar and make it look like the run completed.
    progress.report(processed, len(pdfs), "Stopped" if cancelled else "Done")
    return result
