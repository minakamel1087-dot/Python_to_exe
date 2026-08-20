"""
Clearing a sheet.

Two jobs: wipe any sheet, or wipe the WIR log and put its template row
back.

One change forced by dropping the VBA: the old template restored formulas
in V, X, Z and AB that called a macro (`PathExists`). With no macros in
the workbook those would every one show #NAME?, so those columns are left
empty. Fix Links and Check WIRs both report which paths resolve, which is
what those columns were for.
"""

from __future__ import annotations

from datetime import date, timedelta

from .. import config
from ..findings import RunResult
from ..workbook import LiveWorkbook

TITLE_ANY = "Clear Any Sheet"
TITLE_LOG = "Clear WIR Sheet"

TEMPLATE_LAST_ROW = 100
WIPE_LAST_ROW = 1000
WIPE_LAST_COL = 52              # A:AZ

# The template row, restored after a clear. These are the corrected
# versions: Floor_names for the location, and the Floor_Apartments table
# for the unit list.
DESCRIPTION_FORMULA = (
    '=IFERROR(VLOOKUP(J3,INDIRECT(E3),3,0)&" at "'
    '&IFERROR(TEXTJOIN("-",TRUE,F3:G3,VLOOKUP(H3,Floor_names,3,0),I3), "")'
    '&" As per attached Highlighted Drawings", "")'
    '& IFERROR(IF(D3="",""," (Excluding "&D3&")"),"")'
)
APARTMENTS_FORMULA = '=IFERROR(VLOOKUP(TEXTJOIN("-",TRUE,F3:H3),Floor_Apartments,2,0),"")'
LOCATION_FORMULA = '=IFERROR(TEXTJOIN("-",TRUE,F3:G3,VLOOKUP(H3,Floor_names,3,0),I3), "")'
FORM_FORMULA = '=IFERROR(IF(O3=0,VLOOKUP(J3,INDIRECT(E3),2,FALSE),"Comments_Form"),"")'
CHECK_FORMULA = '=IF(AND(MID(B3,18,3)=F3,MID(B3,29,3)=E3, LEN(B3)=42),"OK","WIR.No Err")'
SITE_LINK_FORMULA = "=HYPERLINK(N3)"
QC_LINK_FORMULA = '=HYPERLINK($W$1&"\\"&B3)'
# The previous revision of this same WIR. Y1 carries the Server/Local
# choice the radio buttons write, and the two names hold the whole prefix
# - so nothing here appends a folder or a separator of its own.
PREV_REV_FORMULA = (
    '=IF(O3=0,"",HYPERLINK(@SWITCH($Y$1,1,ELE_Server_path,2,ELE_Local_path)'
    '&B3&"-R0"&(O3-1)&"-C.pdf"))'
)

# Does the path in the column to the left resolve? PathExists is a VBA
# function in the workbook's Functions module - one of the few pieces of
# VBA still there, and a formula rather than anything this program does.
PATH_EXISTS_COLUMNS = (
    (config.Main.ATT_SITE, "V"),      # U -> V
    (config.Main.ATT_QC, "X"),        # W -> X
    (config.Main.ATT_PREV, "Z"),      # Y -> Z
    (config.Main.ATT_FIRST, "AB"),    # AA -> AB
)


def clear_any(book: LiveWorkbook) -> RunResult:
    """Contents and fills only - borders, fonts, number formats and column
    widths are left alone."""
    result = RunResult(TITLE_ANY)
    sheet = book.excel.ActiveSheet

    target = sheet.Range(
        sheet.Cells(config.FIRST_DATA_ROW, 1),
        sheet.Cells(WIPE_LAST_ROW, WIPE_LAST_COL),
    )
    target.ClearContents()
    target.Interior.ColorIndex = config.COLOUR_NONE

    result.line(f"Cleared rows {config.FIRST_DATA_ROW}-{WIPE_LAST_ROW} on '{sheet.Name}'.")
    return result



def _column_letter(index: int) -> str:
    """1 -> A, 27 -> AA. Needed because the formulas name their source
    column in text, not by number."""
    letters = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def clear_log(book: LiveWorkbook) -> RunResult:
    result = RunResult(TITLE_LOG)
    sheet = book.log_sheet

    target = sheet.Range(
        sheet.Cells(config.FIRST_DATA_ROW, 1),
        sheet.Cells(WIPE_LAST_ROW, WIPE_LAST_COL),
    )
    target.ClearContents()
    target.Interior.ColorIndex = config.COLOUR_NONE

    first, last = config.FIRST_DATA_ROW, TEMPLATE_LAST_ROW

    def fill(column: int, formula: str) -> None:
        sheet.Range(sheet.Cells(first, column), sheet.Cells(last, column)).Formula = formula

    fill(1, CHECK_FORMULA)
    fill(config.Main.DESCRIPTION, DESCRIPTION_FORMULA)
    fill(config.Main.APARTMENTS, APARTMENTS_FORMULA)
    fill(config.Main.LOCATION, LOCATION_FORMULA)
    fill(config.Main.FORM, FORM_FORMULA)
    fill(config.Main.ATT_SITE, SITE_LINK_FORMULA)
    fill(config.Main.ATT_QC, QC_LINK_FORMULA)
    fill(config.Main.ATT_PREV, PREV_REV_FORMULA)

    # Each check sits one column to the right of the path it tests.
    for source, _letter in PATH_EXISTS_COLUMNS:
        fill(source + 1, f"=PathExists({_column_letter(source)}3)")

    submission = sheet.Range(
        sheet.Cells(first, config.Main.DATE_SUBMIT),
        sheet.Cells(last, config.Main.DATE_SUBMIT),
    )
    submission.Value = date.today().strftime("%d-%b-%Y")
    submission.NumberFormat = "dd-mmm-yyyy"

    inspection = sheet.Range(
        sheet.Cells(first, config.Main.DATE_INSPECT),
        sheet.Cells(last, config.Main.DATE_INSPECT),
    )
    inspection.Value = (date.today() + timedelta(days=1)).strftime("%d-%b-%Y")
    inspection.NumberFormat = "dd-mmm-yyyy"

    sheet.Range(
        sheet.Cells(first, config.Main.REV), sheet.Cells(last, config.Main.REV)
    ).Value = 0

    result.line(f"Cleared rows {first}-{WIPE_LAST_ROW} and restored the template to row {last}.")
    result.line("")
    result.line("Columns V, X, Z and AB are left empty - they used to call a")
    result.line("macro. Fix Links and Check WIRs report which paths resolve.")
    return result
