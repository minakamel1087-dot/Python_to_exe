"""
Merging each WIR's documents into one PDF.

Carried over from the previous wir_generate.py, with one deliberate
change: the folder is only deleted when everything in it was a PDF and
therefore made it into the merge. The original deleted the folder
unconditionally, which threw away photos and any other non-PDF
attachment the moment the merge succeeded.
"""

from __future__ import annotations

import os
import shutil

from pypdf import PdfWriter


class MergeResult:
    __slots__ = ("merged_path", "merged_count", "kept_files", "error")

    def __init__(self):
        self.merged_path: str = ""
        self.merged_count: int = 0
        self.kept_files: list[str] = []
        self.error: str = ""

    @property
    def ok(self) -> bool:
        return bool(self.merged_path) and not self.error


def merge_folder(row_folder: str, output_root: str, group_name: str) -> MergeResult:
    """Merges every PDF in `row_folder` into `<group_name>.pdf` directly
    under `output_root`, in filename order so the cover comes first.

    The folder survives whenever the merge does not fully succeed - if it
    failed outright, or if something in there was not a PDF and so could
    not be part of the merge. Either way everything needed to finish the
    job by hand is still sitting where it was left.
    """
    result = MergeResult()

    try:
        entries = sorted(os.listdir(row_folder))
    except OSError as exc:
        result.error = f"cannot read {row_folder}: {exc}"
        return result

    pdfs = [f for f in entries if f.lower().endswith(".pdf")]
    others = [f for f in entries if not f.lower().endswith(".pdf")]

    if not pdfs:
        result.error = "no PDFs to merge"
        result.kept_files = others
        return result

    merged_path = os.path.join(output_root, f"{group_name}.pdf")
    writer = PdfWriter()
    try:
        for name in pdfs:
            writer.append(os.path.join(row_folder, name))
        with open(merged_path, "wb") as handle:
            writer.write(handle)
    except Exception as exc:                       # noqa: BLE001 - reported, not raised
        result.error = str(exc)
        # A half-written merge is worse than none: it looks like a
        # finished document. Remove it, leave the folder exactly as it
        # is, and let the merge be done by hand.
        try:
            if os.path.exists(merged_path):
                os.remove(merged_path)
        except OSError:
            pass
        result.kept_files = entries
        return result
    finally:
        writer.close()

    result.merged_path = merged_path
    result.merged_count = len(pdfs)
    result.kept_files = others

    if others:
        # Something in there is not a PDF - a photo, a DWG, a spreadsheet.
        # It did not make it into the merge, so the folder stays.
        return result

    try:
        shutil.rmtree(row_folder)
    except OSError:
        pass        # merge worked, cleanup did not - not worth failing over

    return result
