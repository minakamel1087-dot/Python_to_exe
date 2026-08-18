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
import threading

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


_share_state: bool | None = None


def share_available(timeout: float = 3.0) -> bool:
    """Is the attachment share actually reachable?

    Off site, or on a machine that never had the drive, every path under
    it fails to resolve. Reporting 52 broken links then is worse than
    useless - it buries the real ones and flags rows that are fine. Ask
    once, and if the answer is no, say so rather than blaming the data.

    Asked on a background thread with a timeout, because a single
    os.path.isdir() against a disconnected mapped drive blocks for about
    twenty seconds while Windows retries the connection - which was the
    whole cost of a run when the share was down. Cached for the process:
    the answer cannot usefully change mid-run.
    """
    global _share_state
    if _share_state is not None:
        return _share_state

    answer: list[bool] = []

    def probe() -> None:
        try:
            answer.append(os.path.isdir(SHARE_ROOT))
        except (OSError, ValueError):
            answer.append(False)

    thread = threading.Thread(target=probe, daemon=True)
    thread.start()
    thread.join(timeout)

    _share_state = bool(answer) and answer[0]
    return _share_state


def folder_reachable(folder: str, timeout: float = 3.0) -> bool:
    """Same trick as share_available, for any folder.

    A disconnected mapped drive makes os.path.isdir block for about twenty
    seconds while Windows retries, so the question is asked on a thread
    and given a deadline. Not cached - unlike the attachment share, this
    is asked once per run and the user may well have just connected.
    """
    if not folder:
        return False

    answer: list[bool] = []

    def probe() -> None:
        try:
            answer.append(os.path.isdir(folder))
        except (OSError, ValueError):
            answer.append(False)

    thread = threading.Thread(target=probe, daemon=True)
    thread.start()
    thread.join(timeout)
    return bool(answer) and answer[0]


def is_on_share(path: str) -> bool:
    return path.upper().startswith(SHARE_ROOT.upper())


def is_network_path(path: str) -> bool:
    """A UNC path, or any drive letter that is not one of this PC's local
    disks - which is what a mis-mapped share looks like.

    Checked before probing when the share is down: a link typed as
    W:\\Common\\... is just as dead as one on Z:, and os.path.exists on
    any disconnected mapped drive blocks for twenty seconds apiece.
    """
    text = unquote(path)
    if text.startswith("\\\\"):
        return True
    if len(text) > 2 and text[1:3] == ":\\":
        return text[0].upper() not in LOCAL_DRIVES
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
