# WIR Tools

A desktop app that works on the **open** workbook. It attaches to the Excel
you already have running and edits the sheet live.

There is no VBA left. The workbook becomes a workbook — data, forms and
reference tables — and this program is the tool.

## Running it

```bat
pip install -r requirements.txt
python main.py
```

`python main.py --check` runs Fix Links → Check Areas → Pre-flight once and
prints the result, with no window. Start there when something misbehaves:
errors print instead of disappearing into a worker thread.

Build the exe with `build.bat` (PyInstaller, one file, no console).

## Why COM and not openpyxl

`openpyxl` and `pandas` read and write the file **on disk**. Against a
workbook that is open, that either fails outright or loses whatever the
user saves next. This uses `pywin32` against the running Excel, so edits
appear in their window immediately.

`GetActiveObject` is used rather than `Dispatch` on purpose — `Dispatch`
would happily start a second, empty Excel and then report that the
workbook is not open.

Excel still renders the PDFs, because `ExportAsFixedFormat` is how an
Excel sheet becomes a PDF and nothing else reproduces the form layout.
That is automation from outside, not code living in the file.

## Layout

```
main.py              window, or --check for the headless run
core/                the logic. Imports nothing from ui.
  config.py          every column, sheet and table name in the workbook
  workbook.py        the live Excel connection, range read/write
  reference.py       the five reference tables, all found by name
  text.py            tokenising and normalising the Area column
  paths.py           attachment link repair
  pdf.py             merging each WIR's documents into one file
  findings.py        Finding, Severity, RunResult
  report.py          the PreFlight sheet
  tasks/
    generate.py        cover + checklist + attachments + merge
    fix_links.py       attachment paths and the register prefix, merged
    check_areas.py     suggestions into column K
    preflight.py       every check, before anything is generated
    import_register.py reload the WIRs sheet from the master register
    import_previous.py approved previous-activity paths into column AA
    history.py         Copy To History, and the Outlook draft
    extract_pdfs.py    copy the visible rows' files into one folder
    clear_sheets.py    wipe a sheet; wipe the log and restore its template
ui/
  theme.py           colours and the stylesheet
  window.py          the window, the worker thread, the buttons
```

`core` never imports from `ui`. That is what lets `--check` exist, and what
would let this run from a scheduled task later.

## One PDF per WIR

Generate renders the cover and the checklist, copies in the attachments,
then merges everything in that WIR's folder into a single
`<WIR No>-R<rev>.pdf` in the output root — carried over from the previous
`wir_generate.py`.

With one deliberate change: the folder is deleted **only when everything
in it was a PDF** and therefore made it into the merge. The original
deleted it unconditionally, which silently threw away photos and any other
non-PDF attachment. Now those keep their folder and the run reports it.

## Reference data stays in the workbook

`Area_Names`, `Activity_Splits`, `WIR_Status_Codes`, `Floor_Apartments`
and `Prev_Activities` are still Excel tables, edited in Excel. They are
looked up **by name**, never by sheet and column — the project data has
already moved sheets twice, and every hardcoded column reference broke
silently when it did.

## What removing the VBA changed

Columns V, X, Z and AB held `=PathExists(...)`, a function that lived in
the workbook's `Functions` module. With no macros those show `#NAME?`, so
Clear WIR Sheet now leaves them empty. Fix Links and Check WIRs both
report which paths resolve, which is what those columns were for.

`Extract PDFs` used to ask you to select a range by hand. It now uses the
Link column for the rows you can see — the same filtered set Copy To
History works on.

## Notes for whoever changes this next

- Read a whole range in one call and write it back in one call. COM
  crosses a process boundary; cell-by-cell loops from Python are slow in a
  way they never were in VBA.
- `text_of()` exists because Excel hands back floats. `101102103.0` must
  not become the string `"101102103.0"`.
- The odd-looking rules in `text.py` are all load-bearing. Each one is
  there because real data broke the previous version — non-breaking
  spaces, unit numbers glued to their prefix, and Excel silently turning
  `101,102,103` into the number `101102103`.
- Dialogs are opened on the GUI thread and the result handed to the
  worker. A file dialog from a worker thread is a crash waiting to happen.
