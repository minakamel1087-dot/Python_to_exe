"""
Copy To History, and the submission email.

Only the rows you can actually see are copied - a filtered log is how you
choose what goes out, so hidden rows are not history.
"""

from __future__ import annotations

import html
from datetime import datetime

from .. import config
from ..findings import RunResult
from ..workbook import LiveWorkbook, text_of

TITLE = "Copy To History"

HISTORY_SHEET = "WIR_History"
TIMESTAMP_COLUMN = 30           # AD
TIMESTAMP_HEADER = "Date Time"
TIMESTAMP_FORMAT = "yyyy-mm-dd hh:mm:ss"

# Excel counts days from 1899-12-30. The timestamps already in the sheet
# are stored that way, so new ones are written as the same kind of number
# rather than as text - a column of dates with strings mixed in cannot be
# sorted or filtered.
EXCEL_EPOCH = datetime(1899, 12, 30)

# Columns that go in the email table: B..J plus O.
EMAIL_COLUMNS = list(range(2, 11)) + [config.Main.REV]

EMAIL_SUBJECT = "WIR Submission"
EMAIL_INTRO = "Please submit the following WIRs"


def run(book: LiveWorkbook, send_email: bool = False, greeting: str = "") -> RunResult:
    result = RunResult(TITLE)
    sheet = book.log_sheet

    if sheet.Name == HISTORY_SHEET:
        return result.fail("You are on the history sheet. Switch to the WIR log first.")

    last_row = book.last_row(sheet, config.Main.WIR_NO)
    if last_row < config.FIRST_DATA_ROW:
        return result.fail(f"No rows found in column B of '{sheet.Name}'.")

    rows = book.block(sheet, config.FIRST_DATA_ROW, last_row, config.Main.LAST_COL)

    visible: list[tuple[int, list]] = []
    for offset, values in enumerate(rows):
        row_no = config.FIRST_DATA_ROW + offset
        if not text_of(values[config.Main.WIR_NO - 1]):
            continue
        if sheet.Rows(row_no).Hidden:
            continue
        visible.append((row_no, values))

    if not visible:
        return result.fail("No visible rows to copy - everything is hidden or filtered out.")

    history = book.ensure_sheet(HISTORY_SHEET)
    if history.FilterMode:
        history.ShowAllData()

    # The history sheet is laid out like the log: row 1 blank, headers on
    # row 2, data from row 3. Reading the header from row 1 would find it
    # empty on a sheet that is perfectly well populated, and appending
    # from row 2 would then write straight over the existing history.
    last_history = book.last_row(history, config.Main.WIR_NO)
    has_header = bool(
        text_of(history.Cells(config.HEADER_ROW, config.Main.WIR_NO).Value2)
    )

    if not has_header:
        headers = book.block(sheet, config.HEADER_ROW, config.HEADER_ROW,
                             config.Main.LAST_COL)
        if headers:
            history.Range(
                history.Cells(config.HEADER_ROW, 1),
                history.Cells(config.HEADER_ROW, config.Main.LAST_COL),
            ).Value = headers
        history.Cells(config.HEADER_ROW, TIMESTAMP_COLUMN).Value = TIMESTAMP_HEADER
        history.Range(
            history.Cells(config.HEADER_ROW, 1),
            history.Cells(config.HEADER_ROW, TIMESTAMP_COLUMN),
        ).Font.Bold = True

    # Never above the first data row, whatever last_row reported - that is
    # the guard that keeps a surprise in the sheet from costing history.
    next_row = max(last_history + 1, config.FIRST_DATA_ROW)

    now = datetime.now()
    stamp = (now - EXCEL_EPOCH).total_seconds() / 86400.0
    block = [values[: config.Main.LAST_COL] + [stamp] for _, values in visible]

    last = next_row + len(block) - 1
    history.Range(
        history.Cells(next_row, 1),
        history.Cells(last, TIMESTAMP_COLUMN),
    ).Value = block

    # Rows are read with Value2, so a date arrives as an Excel serial
    # number. Without a format it lands in history as 45790.
    for column in (config.Main.DATE_SUBMIT, config.Main.DATE_INSPECT):
        history.Range(
            history.Cells(next_row, column), history.Cells(last, column)
        ).NumberFormat = "dd-mmm-yyyy"

    history.Range(
        history.Cells(next_row, TIMESTAMP_COLUMN),
        history.Cells(last, TIMESTAMP_COLUMN),
    ).NumberFormat = TIMESTAMP_FORMAT

    result.line(f"Rows copied:   {len(visible)}")
    result.line(f"Appended to:   rows {next_row}-{last}")
    result.line(f"Stamped:       {now.strftime('%Y-%m-%d %H:%M:%S')}")

    if send_email:
        sent = _open_email(book, sheet, visible, greeting)
        result.line("")
        # `sent` is True, or the reason it failed - and a reason is a
        # non-empty string, so testing it for truth would report every
        # failure as a success.
        result.line("Email drafted in Outlook." if sent is True
                    else f"Email not created: {sent}")

    return result


def _open_email(book: LiveWorkbook, sheet, visible, greeting: str) -> bool | str:
    """Opens a draft in Outlook - never sends it. The table is built from
    the same visible rows that went to history."""
    try:
        import win32com.client

        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
    except Exception as exc:                       # noqa: BLE001
        return str(exc)

    headers = book.block(sheet, 2, 2, config.Main.LAST_COL)
    header_row = headers[0] if headers else []

    parts = [
        "<table border='1' cellpadding='4' cellspacing='0' "
        "style='border-collapse:collapse;font-family:Calibri;font-size:11pt;'>",
        "<tr style='background-color:#00467F;color:white;font-weight:bold;'>",
    ]
    for column in EMAIL_COLUMNS:
        label = text_of(header_row[column - 1]) if column <= len(header_row) else ""
        parts.append(f"<td>{html.escape(label)}</td>")
    parts.append("</tr>")

    for _, values in visible:
        parts.append("<tr>")
        for column in EMAIL_COLUMNS:
            parts.append(f"<td>{html.escape(text_of(values[column - 1]))}</td>")
        parts.append("</tr>")
    parts.append("</table>")

    opening = f"{html.escape(greeting)}<br><br>" if greeting else ""
    mail.Subject = EMAIL_SUBJECT
    mail.HTMLBody = f"{opening}{EMAIL_INTRO}<br><br>{''.join(parts)}<br><br>Best regards<br><br>"
    mail.Display()
    return True
