"""
The live Excel connection.

This attaches to the workbook the user already has open, rather than
opening the file from disk. Everything written here appears in their
window immediately, exactly as the VBA did.

openpyxl and pandas are deliberately absent from this project: they read
and write the file on disk, so against an open workbook they either fail
or lose whatever the user saves next.
"""

from __future__ import annotations

import pythoncom
import win32com.client

from . import config

# Late binding throughout, deliberately. win32com.client.constants and
# gencache.EnsureDispatch depend on a generated type-library cache that
# PyInstaller does not carry into the exe, so the numeric Excel constants
# are spelled out where they are used instead.


class ExcelNotRunning(RuntimeError):
    pass


class WorkbookNotOpen(RuntimeError):
    pass


class WrongSheet(RuntimeError):
    pass


def _values(rng) -> list[list]:
    """A range's values as a list of rows, however Excel chose to hand
    them over. A single cell comes back as a scalar, a single row or
    column as a flat tuple - callers should not have to care.

    Value2, never Value. Value converts a date cell into a COM date,
    which makes pywin32 import win32timezone at that moment - a module
    that is missing unless pywin32's post-install step has run, and one
    PyInstaller cannot see. Value2 hands dates back as plain Excel
    serial numbers, so the dependency never arises. It is faster too.
    """
    raw = rng.Value2
    if raw is None:
        return []
    if not isinstance(raw, tuple):
        return [[raw]]
    if raw and not isinstance(raw[0], tuple):
        return [list(raw)]
    return [list(row) for row in raw]


def text_of(value) -> str:
    """Cell value as trimmed text. Excel hands back floats for anything
    numeric, and 101102103.0 must not become '101102103.0'."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


class LiveWorkbook:
    """Wraps the open workbook and the sheet the user is looking at."""

    def __init__(self, excel, workbook):
        self.excel = excel
        self.wb = workbook
        self._log_sheet = None
        self._sheet_names: set[str] | None = None

    # -- connecting ---------------------------------------------------------

    @classmethod
    def attach(cls, workbook_name: str = config.WORKBOOK_NAME) -> "LiveWorkbook":
        """Finds the running Excel and the already-open workbook.

        GetActiveObject is used rather than Dispatch on purpose: Dispatch
        would happily start a second, empty Excel and then report that the
        workbook is not open.
        """
        pythoncom.CoInitialize()
        try:
            excel = win32com.client.GetActiveObject("Excel.Application")
        except pythoncom.com_error as exc:
            raise ExcelNotRunning(
                "Excel is not running. Open the WIR workbook first, then try again."
            ) from exc

        wanted = workbook_name.lower()
        for book in excel.Workbooks:
            if book.Name.lower() == wanted:
                return cls(excel, book)

        open_names = ", ".join(b.Name for b in excel.Workbooks) or "nothing"
        raise WorkbookNotOpen(
            f"'{workbook_name}' is not open in Excel. Currently open: {open_names}."
        )

    # -- sheets -------------------------------------------------------------

    @property
    def log_sheet(self):
        """The sheet the user was on when the run started.

        Resolved once and held for the life of the connection. Re-reading
        ActiveSheet per task means a run that spans several seconds breaks
        if the user clicks another tab - or, far worse, writes to it.
        """
        if self._log_sheet is not None:
            return self._log_sheet

        sheet = self.excel.ActiveSheet
        if sheet.Name in config.NON_LOG_SHEETS:
            raise WrongSheet(
                f"'{sheet.Name}' is not a WIR log sheet. "
                "Switch to your log sheet in Excel, then run this again."
            )
        self._log_sheet = sheet
        return sheet

    def sheet(self, name: str):
        for sheet in self.wb.Worksheets:
            if sheet.Name == name:
                return sheet
        return None

    def sheet_exists(self, name: str) -> bool:
        """Answered from a cached set of names.

        Pre-flight asks this once per row for the checklist in column P.
        Walking Worksheets each time is 60 rows x 30 sheets of COM calls
        for a question whose answer cannot change during a run.
        """
        if self._sheet_names is None:
            self._sheet_names = {s.Name.lower() for s in self.wb.Worksheets}
        return name.lower() in self._sheet_names

    def ensure_sheet(self, name: str):
        found = self.sheet(name)
        if found is not None:
            return found
        created = self.wb.Worksheets.Add(After=self.wb.Worksheets(self.wb.Worksheets.Count))
        created.Name = name
        return created

    # -- reading ------------------------------------------------------------

    def last_row(self, sheet, column: int) -> int:
        return sheet.Cells(sheet.Rows.Count, column).End(-4162).Row   # xlUp

    def block(self, sheet, first_row: int, last_row: int, last_col: int) -> list[list]:
        """A whole rectangle in one call. Reading cell by cell across the
        process boundary is what makes COM feel slow; this is the reason
        it does not."""
        if last_row < first_row:
            return []
        rng = sheet.Range(sheet.Cells(first_row, 1), sheet.Cells(last_row, last_col))
        return _values(rng)

    def table_values(self, table_name: str) -> list[list]:
        """An Excel Table's data rows, found by name anywhere in the
        workbook. Table names are unique workbook-wide, so whichever sheet
        it currently lives on does not matter - which is the whole point of
        looking it up this way."""
        for sheet in self.wb.Worksheets:
            for table in sheet.ListObjects:
                if table.Name.lower() == table_name.lower():
                    body = table.DataBodyRange
                    return _values(body) if body is not None else []
        return []

    def named_values(self, name: str) -> list[list]:
        """A defined name's values, whether it is scoped to the workbook or
        to a single sheet."""
        try:
            return _values(self.wb.Names(name).RefersToRange)
        except pythoncom.com_error:
            pass
        for sheet in self.wb.Worksheets:
            try:
                return _values(sheet.Names(name).RefersToRange)
            except pythoncom.com_error:
                continue
        return []

    # -- writing ------------------------------------------------------------

    def write_column(self, sheet, column: int, first_row: int, values: list) -> None:
        """Writes a whole column in one call. One round trip instead of
        several hundred."""
        if not values:
            return
        rng = sheet.Range(
            sheet.Cells(first_row, column),
            sheet.Cells(first_row + len(values) - 1, column),
        )
        rng.Value = [[v] for v in values]

    def clear_fill(self, sheet, column: int, first_row: int, last_row: int) -> None:
        if last_row < first_row:
            return
        rng = sheet.Range(sheet.Cells(first_row, column), sheet.Cells(last_row, column))
        rng.Interior.ColorIndex = config.COLOUR_NONE

    def fill(self, sheet, row: int, column: int, colour: int) -> None:
        sheet.Cells(row, column).Interior.Color = colour

    # -- run state ----------------------------------------------------------

    def __enter__(self) -> "LiveWorkbook":
        self._screen = self.excel.ScreenUpdating
        self.excel.ScreenUpdating = False
        return self

    def __exit__(self, *exc) -> None:
        self.excel.ScreenUpdating = self._screen
        try:
            self.excel.StatusBar = False
        except pythoncom.com_error:
            pass
        return False
