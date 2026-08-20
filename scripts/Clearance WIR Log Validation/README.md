# Clearance WIR Log Validation

Validates the ELE WIR log and the MEP Clearance log against the project's own
registers, cross-matches them so open clearances can be closed out, and builds
a progress tracker.

A Windows desktop app. Everything lives in one workbook, so there is one file to
choose and nothing to configure per run.

## Layout on disk

```
Clearance WIR Log Validation/             <- the workbook lives here
    WIR_Clearance Log Validation.xlsm     <- the one input
    Clearance WIR Match/                  <- the four outputs, created on first run
    ClearanceWIRLogValidation/            <- the portable app; copy this folder anywhere
        ClearanceWIRLogValidation.exe
        _internal/
    Clearance WIR Log Validation/         <- this repo
```

The app is a folder rather than a single exe, so deploying it is a copy — no
rebuild. It finds the workbook beside itself or a level or two up, and writes
its output beside the workbook, so results land next to the data they describe.

## The workbook

`WIR_Clearance Log Validation.xlsm` carries everything:

| Sheet | What it is |
|---|---|
| `WIRs` | the WIR log — headers on row 3 |
| `Clearance` | the clearance log — headers on row 3 |
| `Project-Des.` | **the single source of reference data**: buildings, floors, the 1,081-unit apartment register, WIR status letters, and the area vocabulary with its spelling variants |
| `ClearanceWIRMap`, `Constants`, `ErrorCodes` | copies of the app's configuration; the app reads `reference/` in preference to these |

**Change a rule by editing `Project-Des.`** Nothing about the project is
duplicated in this repo, so the two cannot drift.

## What it produces

Into `Clearance WIR Match/`, beside the workbook:

| Workbook | |
|---|---|
| `ELE_WIR_Validated.xlsx` | doc 01 — Data / Exceptions / Mapping / Dashboard |
| `MEP_Clearance_Validated.xlsx` | doc 02 — same four sheets |
| `Clearance_WIR_Match.xlsx` | doc 04 — Matches / Gaps / Summary |
| `Clearance_Progress.xlsx` | doc 03 — the wide tracker |

### Reading the match report

`Matches` holds **one row per matched WIR**, so a clearance covered by three
WIRs occupies three rows and every path cell holds exactly one path — which is
what makes the two `HYPERLINK` columns work. Clearance-level values repeat down
the group, so the sheet stays a flat table that filters and pivots normally.
The `Summary` sheet and the run headline count *clearances*, not rows.

The `Last` column carries an `L` on the newest revision of each clearance
number. It is worked out across every validated clearance, not just the ones
the report checks — so a clearance whose newest revision was approved shows its
open earlier revisions with no `L`, which is the signal that they are
superseded.

## Importing raw logs

The two Import buttons are the Python port of the workbook's own macros. Each
reads a raw log — off the server or off disk — and replaces the `WIRs` or
`Clearance` sheet with its contents. **Close the workbook in Excel first.**

Reading uses openpyxl; writing goes through Excel itself, because the master
carries macros, tables and external links that a full openpyxl round-trip can
silently drop. Excel touches only the cells asked for.

Three things differ from the macros, all deliberate:

- **ELE only.** The macro also pulled `ELV`, which is what put a second header
  row in the middle of the data — the stray row reading `Dep.` — and added 400
  rows the activity mapping has nothing to say about.
- **The whole sheet is cleared first.** The VBA cleared `A:AD`, so anything
  further right survived and old data mixed into new. Contents go; formats and
  column widths stay, so imported serials still render as dates and times.
- `ImportCleranceLog` was written over a copy of `ImportWIRData` and its column
  positions were never moved. Its `AA` column read the ELEC link and came out as
  `(\\192.168…pdf)`; `AB` joined Description-Package-Building-AREA; `AC` held
  the Level rather than the Area; and `AD` had a header with nothing behind it.
  All four now use the clearance's own columns. See `engine/importer.py`.

## Running from source

```bash
pip install -r requirements-dev.txt
```

```bash
python -m pytest
```

```bash
python desktop/app.py
```

Tests that assert real figures need the workbook and skip without it. It is
looked for beside the repo and one level up; set `GHAF_WORKBOOK` to point
elsewhere. It is deliberately not committed — it carries internal network paths
and commercial data.

## Building the app

```powershell
.\build.ps1
```

Runs the tests, builds the folder, runs its own `--selftest`, and only then
copies it beside the workbook. Add `-NoDeploy` to build without deploying.

The self-check is the part that matters. It generates a synthetic workbook in
memory, runs all four programs against it, drives Excel, and performs a real
import from a source carrying a date and a time — so it needs no project data
and it catches the failures that only appear once packaged. Two shipped before
it did that: a pywin32 module PyInstaller could not see, and a `datetime.time`
COM refuses to marshal. Both would now fail the build.

## Provenance

The apartment register holds **1,081** units. Three blocks present in an early
source do not exist and are excluded: `C-002 5F 511–515`, `C-009 5F 511–515`,
and `A2-005 4F 406–407`. Reconcile against two independent sources before
trusting a regenerated count — top floors are prone to over-extension by
copy-paste from the floor below.

Basement apartments exist only where the register says so. "Type C buildings
have basements" holds for P08 and is false for P09, so eligibility is always a
register lookup, never inferred from the building code.
