# WIR Tools

A desktop app that works on the **open** workbook. It attaches to the Excel
you already have running and edits the sheet live — same as the VBA did,
just from outside.

## Running it

```bat
pip install -r requirements.txt
python main.py
```

`python main.py --check` runs Fix Links → Check Areas → Pre-flight once and
prints the result, with no window. That form is how the logic gets tested.

Build the exe with `build.bat` (PyInstaller, one file, no console).

## Why COM and not openpyxl

`openpyxl` and `pandas` read and write the file **on disk**. Against a
workbook that is open, that either fails outright or loses whatever the
user saves next. This project uses `pywin32` and talks to the running
Excel, so edits appear in their window immediately and their undo history
behaves the way it does for any other macro.

`GetActiveObject` is used rather than `Dispatch` on purpose — `Dispatch`
would happily start a second, empty Excel and then report that the
workbook is not open.

## Layout

```
main.py              entry point: window, or --check for the headless run
core/                the logic. Imports nothing from ui.
  config.py          every column, sheet and table name in the workbook
  workbook.py        the live Excel connection and range read/write
  reference.py       Prev_Activities, Floor_Apartments, Activity_Splits,
                     WIR_Status_Codes, Area_Names - all found by name
  text.py            tokenising and normalising what people type in Area
  paths.py           attachment link repair
  findings.py        Finding, Severity, RunResult
  report.py          the PreFlight sheet
  tasks/
    fix_links.py     Fix Attachment Links + Fix WIR Path Prefix, merged
    check_areas.py   suggestions into column K
    preflight.py     every check, before anything is generated
ui/
  theme.py           colours and the stylesheet
  window.py          the window, the worker thread, the buttons
```

`core` never imports from `ui`. That is what lets `--check` exist, and what
would let this run from a scheduled task later.

## What is ported and what still calls VBA

Ported to Python — all the logic that had grown complicated and deserved
tests:

- Fix Links, Check Areas, Pre-flight
- the tokeniser, the register/precedence/split matching, the report

Still calling the workbook's macros, through `Application.Run`:

- Generate WIRs, Copy To History, Extract PDFs, Import WIR Data,
  Import Previous WIR Paths, Clear WIR Sheet, Clear Any Sheet

Those are Excel's own jobs — rendering sheets to PDF, driving Outlook,
restoring template formulas. Reimplementing them in Python would mean
reimplementing Excel. The buttons look identical either way.

## Reference data stays in the workbook

`Area_Names`, `Activity_Splits`, `WIR_Status_Codes`, `Floor_Apartments`
and `Prev_Activities` are still Excel tables, edited in Excel. They are
looked up **by name**, never by sheet and column — the project data has
already moved sheets twice, and every hardcoded column reference broke
silently when it did.

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
