"""
Server and local paths, and which of the two the program is using.

The pairs live in the workbook, on the **Paths** sheet:

    Item | Server                                  | Local
    ELE  | \\\\192.168.225.6\\...\\Commented WIR\\ELE&ELV\\ | D:\\i125\\Logs\\WIR\\ELE&ELV\\extracted\\
    ELV  | ...                                     | ...

They are in the workbook rather than in the program because each person
takes a copy of the sheet and the .exe and points the local column at
whatever their own PC uses. Nothing here is hard-coded.

The mode - Server or Local - is a number in **Main!Y1**, so the sheet's
own formulas can follow it:

    =IF(O3=0,"",HYPERLINK(@SWITCH($Y$1,1,ELE_Server_path,2,ELE_Local_path)
                          &B3&"-R0"&(O3-1)&"-C.pdf"))

The program writes that cell when the radio buttons change, and resets it
to Server every time it starts.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from . import config
from .workbook import LiveWorkbook, text_of

SERVER = 1
LOCAL = 2


def _with_separator(folder: str) -> str:
    r"""A prefix has to end in a separator to join onto a file name.

    The Local column is easy to type without the trailing backslash, and
    "…\extracted" + "UAE044…" gives "…\extractedUAE044…" - a path that
    cannot exist, reported as a missing file. Fixed here so the program
    is right either way. **Excel formulas are not protected by this**:
    a named range used as `Local_path & B3` needs the backslash typed in.
    """
    folder = folder.strip().strip('"').rstrip()
    if not folder:
        return ""
    return folder if folder.endswith(("\\", "/")) else folder + "\\"


@dataclass
class PathPair:
    item: str
    server: str
    local: str

    @property
    def usable(self) -> bool:
        return bool(self.server and self.local)


@dataclass
class PathMap:
    pairs: list[PathPair] = field(default_factory=list)
    missing_local: list[str] = field(default_factory=list)

    def to_local(self, path: str) -> tuple[str, str]:
        """(path, warning). A path under a server prefix comes back under
        the matching local one. Anything else is returned untouched.

        A pair with no local twin yields the original path and a warning -
        carrying on with the server path is more useful than refusing,
        and the warning says the mirror is incomplete.
        """
        if not path:
            return path, ""

        lowered = path.lower()
        for pair in self.pairs:
            if not pair.server or not lowered.startswith(pair.server.lower()):
                continue
            if not pair.local:
                return path, (f"{pair.item}: no local path in the Paths sheet, "
                              f"using the server path")
            return pair.local + path[len(pair.server):], ""

        return path, ""

    def server_of(self, path: str) -> tuple[str, str]:
        """The reverse, for completeness - a local path back to its
        server twin."""
        if not path:
            return path, ""
        lowered = path.lower()
        for pair in self.pairs:
            if pair.local and lowered.startswith(pair.local.lower()):
                return pair.server + path[len(pair.local):], ""
        return path, ""


def load(book: LiveWorkbook) -> PathMap:
    """The Paths sheet, or an empty map when it is not there.

    An absent sheet is not an error: the program worked without one until
    now, and every caller treats an empty map as "leave paths alone".
    """
    result = PathMap()
    sheet = book.sheet(config.PATHS_SHEET)
    if sheet is None:
        return result

    last = book.last_row(sheet, config.Paths.ITEM)
    for row in range(config.Paths.FIRST_ROW, last + 1):
        item = text_of(sheet.Cells(row, config.Paths.ITEM).Value2).strip()
        server = _with_separator(text_of(sheet.Cells(row, config.Paths.SERVER).Value2))
        local = _with_separator(text_of(sheet.Cells(row, config.Paths.LOCAL).Value2))
        if not item and not server:
            continue
        result.pairs.append(PathPair(item, server, local))
        if server and not local:
            result.missing_local.append(item or server)

    # Longest server prefix first, so a more specific row wins over a
    # shorter one that happens to be its parent folder.
    result.pairs.sort(key=lambda p: len(p.server), reverse=True)
    return result


# --- which mode the program is in -----------------------------------------

def _mode_sheet(book: LiveWorkbook):
    """The one sheet holding the mode cell.

    By name, never "the active sheet". The SWITCH() formulas that read Y1
    live on one sheet, and taking the active tab instead wrote the value
    onto whichever tab happened to be showing - reporting success while
    the formulas saw nothing.

    Falls back to the active log sheet only if that named sheet is absent.
    """
    sheet = book.sheet(config.MODE_SHEET)
    if sheet is not None:
        return sheet
    try:
        return book.log_sheet
    except Exception:                              # noqa: BLE001
        return None


def read_mode(book: LiveWorkbook) -> int:
    """Server unless the cell plainly says Local."""
    sheet = _mode_sheet(book)
    if sheet is None:
        return SERVER
    try:
        raw = sheet.Cells(config.MODE_ROW, config.MODE_COL).Value2
        return LOCAL if int(float(raw)) == LOCAL else SERVER
    except (TypeError, ValueError, Exception):     # noqa: BLE001
        return SERVER


def write_mode(book: LiveWorkbook, mode: int) -> bool:
    """Put it where the sheet formulas read it. True if it landed."""
    sheet = _mode_sheet(book)
    if sheet is None:
        return False
    try:
        sheet.Cells(config.MODE_ROW, config.MODE_COL).Value = int(mode)
        return True
    except Exception:                              # noqa: BLE001
        return False                               # a locked sheet is not fatal
