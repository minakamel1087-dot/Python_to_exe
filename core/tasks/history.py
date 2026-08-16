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

    last_history = book.last_row(history, config.Main.WIR_NO)
    if last_history < 1 or not text_of(history.Cells(1, config.Main.WIR_NO).Value):
        headers = book.block(sheet, 2, 2, config.Main.LAST_COL)
        if headers:
            history.Range(
                history.Cells(1, 1), history.Cells(1, config.Main.LAST_COL)
            ).Value = headers
        history.Cells(1, TIMESTAMP_COLUMN).Value = "History_Timestamp"
        history.Range(history.Cells(1, 1), history.Cells(1, TIMESTAMP_COLUMN)).Font.Bold = True
        next_row = 2
    else:
        next_row = last_history + 1

    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    block = [values[: config.Main.LAST_COL] + [stamp] for _, values in visible]

    history.Range(
        history.Cells(next_row, 1),
        history.Cells(next_row + len(block) - 1, TIMESTAMP_COLUMN),
    ).Value = block

    result.line(f"Rows copied:   {len(visible)}")
    result.line(f"Appended from: row {next_row}")

    if send_email:
        sent = _open_email(book, sheet, visible, greeting)
        result.line("")
        result.line("Email drafted in Outlook." if sent else f"Email not created: {sent!r}")

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
