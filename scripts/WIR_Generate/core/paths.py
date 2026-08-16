"""
Attachment links arrive from other people's machines and are wrong in
three predictable ways:

    "Z:\\WIR request\\..."   wrapped in quotes (Explorer's "Copy as path")
    S:\\WIR request\\...     the share mapped to a different drive letter
    Z:\\WIR request\\...     the \\Common\\ level missing

Repair is deliberately conservative: a local path is never touched, and a
corrected path is only ever used when the files are actually there
afterwards.
"""

from __future__ import annotations

import os

from .config import LOCAL_DRIVES, SHARE_ROOT
from .text import normalize_spaces

_QUOTES = "\"'\u201c\u201d"

# A cell may hold several paths. These are the separators the copy step
# has always understood; keeping them identical means a cell that worked
# before still works.
_SPLITTERS = (";", "\n", "\r")


def split_paths(raw: str) -> list[str]:
    """The individual paths inside one cell, in order, untrimmed of their
    original spelling."""
    if not raw:
        return []
    work = raw
    for sep in _SPLITTERS:
        work = work.replace(sep, "|")
    return [p.strip() for p in work.split("|") if p.strip()]


def unquote(path: str) -> str:
    return normalize_spaces(path).strip().strip(_QUOTES).strip()


def trim_slashes(path: str) -> str:
    while len(path) > 3 and path.endswith("\\"):
        path = path[:-1]
    return path


def repair(path: str) -> str:
    """Quotes removed, drive forced to Z:, and everything hung off
    Z:\\Common\\.

    A local drive (C: D: E: F: G: H:) comes back untouched - those are
    real local folders, not a mis-mapped share. UNC paths are left alone
    too.
    """
    text = trim_slashes(unquote(path))
    if len(text) < 4 or text[1:3] != ":\\":
        return text

    if text[0].upper() in LOCAL_DRIVES:
        return text

    rest = text[3:]
    if rest[:7].upper() == "COMMON\\":
        return SHARE_ROOT + rest[7:]
    return SHARE_ROOT + rest


def exists(path: str) -> bool:
    """A file or a folder - either is a valid attachment target."""
    if not path:
        return False
    try:
        return os.path.exists(path)
    except (OSError, ValueError):
        return False


def repair_cell(raw: str) -> tuple[str, bool, list[str]]:
    """Repairs every path inside one cell, in place, so whatever
    separators the cell used survive untouched.

    Returns (new_text, changed, still_broken).
    """
    if not raw or not raw.strip():
        return raw, False, []

    new_text = raw
    changed = False
    broken: list[str] = []

    for original in split_paths(raw):
        bare = trim_slashes(unquote(original))

        if exists(bare):
            # The path was fine; only the quotes were in the way.
            if bare != original:
                new_text = new_text.replace(original, bare)
                changed = True
            continue

        fixed = repair(original)
        if fixed != original and exists(fixed):
            new_text = new_text.replace(original, fixed)
            changed = True
        else:
            # Not found even after correcting - leave it exactly as typed.
            broken.append(bare or original)

    return new_text, changed, broken
