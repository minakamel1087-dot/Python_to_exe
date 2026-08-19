# WIR Generate Tools — Technical Reference

Complete specification of the program: what it does, the rules it applies,
the workbook contract it depends on, and why each decision was made.

Written to be enough to rebuild the program from scratch, or to review it
without reading every line.

---

## 1. What this is

A Windows desktop application that automates the Work Inspection Request
(WIR) process for the i125 / Ghaf Woods project. It replaces a set of
Excel VBA macros that previously lived inside
`i125-WIR Cover Generate.xlsm`.

It does four kinds of work:

| Kind | Tasks |
|---|---|
| Produce documents | Generate WIRs |
| Repair data | Fix Attachment Links, Fix WIR Path Prefix, Check Areas |
| Validate before submission | Check WIRs Before Generate |
| Move data around | Import WIR Data, Import Previous WIR Paths, Copy To History, Extract PDFs, Clear sheets |

### 1.1 The central design decision

**The program edits the workbook the user already has open.** It does not
open the file from disk, and it does not ask for a file path.

This is done with COM automation (`pywin32`) against the running Excel
instance. Every write appears in the user's window immediately.

`openpyxl` and `pandas` are deliberately **not used anywhere**. They read
and write the file on disk; against an open workbook that either fails or
is silently overwritten when the user next saves.

```python
excel = win32com.client.GetActiveObject("Excel.Application")
```

`GetActiveObject`, not `Dispatch`. `Dispatch` will happily start a second,
empty Excel instance and then report that the workbook is not open.

### 1.2 Excel is still required

Excel renders the PDFs. `Worksheet.ExportAsFixedFormat` is the only thing
that reproduces the WIR form layout — its fonts, borders, merged cells,
print areas and embedded logos. Reimplementing that means reimplementing
Excel.

That is automation from outside the file, not code living inside it. The
workbook contains **no VBA**.

---

## 2. Runtime and dependencies

| Component | Purpose | Notes |
|---|---|---|
| Python 3.11+ | runtime | tested on 3.13.15 |
| `pywin32` >= 308 | COM to Excel and Outlook | 3.13 needs >= 308 |
| `pypdf` >= 4.3 | merging each WIR's PDFs | |
| `PySide6` >= 6.8 | the window | not needed for `--check` / `--probe` |
| Microsoft Excel | rendering sheets to PDF | must be running with the workbook open |
| Microsoft Outlook | optional, submission email draft | only for Copy To History |

Version floors matter: PySide6 6.7.2 and pywin32 306 have **no wheels for
Python 3.13**. Pinning them exactly makes the project uninstallable on a
current interpreter.

### 2.1 Entry points

```
python main.py            open the window
python main.py --check    Fix Attachment Links -> Check Areas -> Pre-flight, printed
python main.py --probe    read-only: attach, load reference tables, count rows
```

`--probe` writes nothing. It is the first thing to run on a new machine:
it proves the COM attach, the workbook lookup, table resolution by name,
and `Value2` reads without touching data.

### 2.2 Two startup shims in `main.py`

Both run before anything else is imported.

**`_add_project_to_path()`** — an embeddable Python builds `sys.path`
from its `._pth` file and does **not** add the script's own directory, so
`import core` fails. Normal installs already have this entry.

**`_register_pywin32_dlls()`** — pywin32 keeps `pythoncom` and
`pywintypes` in `site-packages/pywin32_system32` and relies on a
post-install step to make them findable. That step cannot run in an
embeddable Python, so the folder is registered explicitly with
`os.add_dll_directory`.

---

## 3. Project structure

```
main.py                 entry point and the two startup shims
requirements.txt
build.bat               PyInstaller one-file build

core/                   all logic. Imports nothing from ui.
  config.py             every column, sheet and table name in the workbook
  workbook.py           the live Excel connection, range read/write
  reference.py          the five reference tables, loaded once per run
  text.py               tokenising and normalising the Area column
  paths.py              attachment link repair and share detection
  pdf.py                merging each WIR's documents into one file
  findings.py           Finding, Severity, RunResult
  report.py             writes the PreFlight sheet
  tasks/
    generate.py         cover + checklist + attachments + merge
    fix_attachments.py  repair the log's attachment paths
    fix_prefix.py       repoint the register's paths to the local folder
    check_areas.py      suggestions and notes into column K
    preflight.py        every validation, before anything is generated
    import_register.py  reload the WIRs sheet from the master register
    import_previous.py  approved previous-activity paths into column AA
    history.py          Copy To History, and the Outlook draft
    clear_sheets.py     wipe a sheet; wipe the log and restore its template

ui/
  theme.py              colours and the Qt stylesheet
  window.py             the window, the worker thread, the buttons
```

**`core` never imports from `ui`.** That is what allows `--check` and
`--probe` to exist, what makes the logic testable, and what would let this
run from a scheduled task.

---

## 4. The workbook contract

Everything the program assumes about the workbook. Change any of this and
`core/config.py` is the only file that needs editing.

### 4.1 The log sheet

Whichever sheet is **active** when a task starts. Guarded against a list
of known non-log sheets (`WIR-Form`, `PreFlight`, `WIRs`, `WIR_History`,
`Project-Des.`, `ELE_Data`, `PUB_Data`, `MacroLog`, …).

Resolved **once per connection and cached**. Re-reading `ActiveSheet` per
task means a run spanning several seconds breaks if the user clicks
another tab — or worse, writes to it.

Headers on row 2. Data from row 3.

| Col | # | Meaning | Written by |
|---|---|---|---|
| A | 1 | Sr / number check formula | Clear WIR Sheet |
| B | 2 | Submittal Reference No. (42 chars) | user |
| C | 3 | Description of Submission | formula |
| D | 4 | Exclusion | user |
| E | 5 | Dep. (trade: ELE, ELV, PUB…) | user |
| F | 6 | Package (P08, P09, P13) | user |
| G | 7 | Building | user |
| H | 8 | Floor | user |
| I | 9 | **Area** | user |
| J | 10 | Activity | user |
| K | 11 | **Suggested area / notes** | **Check Areas** |
| L | 12 | Floor apartments | formula |
| M | 13 | LOCATION | formula |
| N | 14 | **Link** — the typed attachment path | user, **Fix Attachment Links** |
| O | 15 | Rev | user |
| P | 16 | Form — the checklist sheet name | formula |
| S | 19 | Date of submission | Clear WIR Sheet |
| T | 20 | Date of inspection | Clear WIR Sheet |
| U | 21 | Site Attachments — `=HYPERLINK(N3)` | formula |
| W | 23 | QC Attachment — `=HYPERLINK($W$1&"\"&B3)` | formula |
| Y | 25 | Previous REV | formula |
| AA | 27 | **Previous WIR paths** | **Import Previous WIR Paths** |
| AC | 29 | Status (`Done` when generated) | **Generate WIRs** |

The WIR number encodes the package at characters 18–20 and the trade at
29–31, and is 42 characters long. Pre-flight checks all three.

### 4.2 The register — `WIRs` sheet

Headers on row 3, data from row 4. Rebuilt wholesale by Import WIR Data.

| Col | # | Meaning |
|---|---|---|
| B | 2 | WIR number |
| F | 6 | Package |
| G | 7 | Building |
| H | 8 | Floor |
| I | 9 | Area |
| J | 10 | Activity |
| N | 14 | Path on the network share |
| O | 15 | Rev |
| P | 16 | Status code |
| AA | 27 | WIR No and Status (derived) |
| AB | 28 | Code (derived) |
| AC | 29 | Area (derived) |
| AD | 30 | **Local Path** (derived) |

**N versus AD matters.** N is where the register stores the commented WIR
on the share; AD is the same file under the local extract folder. AD is
the copy that actually opens, and it is what Import Previous WIR Paths
copies into the log.

### 4.3 Reference tables — found by name, never by position

The project data has already been split across sheets twice. Every
hardcoded column reference broke silently when it happened. All five are
resolved by Excel Table name or defined name, anywhere in the workbook.

#### `Prev_Activities` (Excel Table)

`Current Activity | Previous Activity 1..n`

Which activities must be approved before this one. Currently 33
activities.

#### `Floor_Apartments` (Excel Table)

`Code | Apartment | Package | Building | Floor | Apt01 … Apt15`

Code is `P08-A1-008-GF`. The `Apt01..Apt15` columns are preferred — one
unit per cell, nothing to parse — with the free-text `Apartment` column as
fallback. Currently 123 floor codes.

#### `Activity_Splits` (defined name)

`Split Activity | Sub-Activity 1..n`

Activities raised as two or more sub-activities instead of one:

```
Wires_Pulling      -> Wires_H.L + Wire_PWR_Pulling
DB&ONU_Enclosure   -> DB_Enclosure + ONU_Enclosure
```

When every sub is approved the parent counts as done, and **no WIR for
the parent will ever exist**. Loaded in both directions: parent → subs,
and sub → parents.

#### `WIR_Status_Codes` (defined name)

`Code | Meaning | Counts as approved`

| Code | Meaning | Approved |
|---|---|---|
| B | Approved | **Y** |
| C | Rejected | N |
| D | Rejected | N |
| R | Under review | N |
| S | Terminated | N |
| T | Terminated | N |

The "Counts as approved" column drives **every** approval decision in the
program. An unmapped code renders as `Status X` rather than going blank.

#### `Area_Names` (defined name)

`Area | Floors | Variant 1..n`

- **Area** — how the area should be written. This is what Check Areas
  writes into the suggestion.
- **Floors** — comma list, or blank for anywhere. Only used to reject an
  area on a floor it cannot be on.
- **Variants** — spellings seen in the log that mean the same place.

Floor-restricted entries currently:

```
Main Electrical Room       BS
Essential Electrical Room  BS
Main Telephone Room        BS
Sump Pit                   BS
Top Roof                   RF
```

**Main Electrical Room and Electrical Room are different rooms**, as are
Main Telephone Room and Telephone Room. Generic room-matching libraries
collapse `Main/Essential/Electrical` into one family; doing that here
would hide a real mistake. `MTR` means Main Telephone Room on this
project — confirmed from the register, where all 8 MTR rows are in the
basement, matching Main Telephone Room and unlike Water Meter Room which
spans every floor.

---

## 5. Text processing — `core/text.py`

Every rule here exists because real data broke a previous version. None
of it is defensive programming for its own sake.

### 5.1 `normalize_spaces`

Non-breaking spaces (U+00A0), tabs and line breaks become plain spaces,
and runs of spaces collapse.

Cells pasted from Word, Outlook or a web page carry non-breaking spaces.
`str.strip()` does not remove them. A register entry reading
`Apartment\u00a0201` tokenised to `\u00a0201`, matched nothing, and made a
whole activity look unapproved with no visible reason why.

### 5.2 `tokenize`

Splits an Area value into parts **exactly as typed** — the report quotes
these back, so `Garbage Room` must not become `Garbage`.

1. Normalise spaces, strip `(` and `)`.
2. Split on `,` `;` `/` `&` and the word ` and `.
3. Drop empties and the literal `MISSING`.
4. Expand number runs (5.3).
5. Inherit names (5.4).

Brackets are stripped because `(Staircase 1, 2 and Lift Core)` otherwise
split into `(Staircase 1` and `2 and Lift Core )`.

**Hyphens are deliberately not separators.** They are load-bearing inside
`Part-5`, `Mock-up`, `Non-Tower`, `Staircase-1` and `A1 - 004`. Splitting
on them would reduce every `Part-N` to a bare `PART` token and create
false duplicate matches. Only 30 of 1,788 register areas contain a hyphen,
so the cost of not splitting is small and the cost of splitting is large.

### 5.3 `_expand_number_runs`

`301 302 303 304` and `Apartments 201 202` arrive as one token because
nobody typed the commas.

A token is split **only when every part is a unit code**. `G10 Kitchens`
contains a room name and stays whole.

### 5.4 `_inherit_names`

`Staircase 1, 2` means staircase one and staircase two.

A bare number following a `<name> <number>` token takes that name. Applies
only when the previous token has a non-numeric stem and a space before its
number, so `301, 302` and `G01, G02` are untouched — those are already
complete on their own.

### 5.5 `match_key`

The comparable form of one token: unit-word prefix removed, trailing
`ROOM`/`ROOMS` removed, upper-cased.

**Only the leading run of letters is treated as the prefix, never the
first whole word.** The log contains both `Apartment 201` and
`Apartment201`; dropping the whole word would swallow the number in the
second, so the two would never match each other.

### 5.6 `is_unit_word` — deliberately fuzzy

A leading word counts as meaning "apartment" if it starts with `APART` or
`FLAT`, is one of `APT UNIT NO PACKAGE BUILDING`, or is **within two edits
of `APARTMENT`/`APARTMENTS`**.

The log contains `Apartmants`, `Aparments` and `Aartments`. A strict test
leaves those looking like room names, and the number inside them never
gets checked against the register.

### 5.7 `is_unit_for_floor`

Does a code fall in the floor's own numbering band?

```
GF        G01 .. G99      (G + 2 digits)
1F .. 10F floor digits + 2 more   (1F -> 1xx, 10F -> 10xx)
anything else             no band
```

Used to stop a number being *relabelled* as an apartment when it plainly
is not one.

### 5.8 `split_glued_numbers`

Excel reads `101,102,103` as a number with thousand separators and stores
**101102103**. The commas are destroyed at typing time, before any code
sees the cell.

Recovery: split into 3-digit codes, accept **only if every chunk lands in
the floor's band**. `101102103` on 1F → `101, 102, 103`. Any chunk that
does not match means nothing is returned, rather than a guess.

### 5.9 `area_key` and `nearest`

`area_key` is the comparison form of an area *name*: letters and digits
only, upper case, trailing `S` dropped so `Store Rooms` and `Store Room`
are one thing.

`nearest` finds the closest known name by Levenshtein distance, capped at:

```
_MAX_EDITS      = 2
_MAX_LENGTH_GAP = 2
```

The length guard is the important half. Without it,
`Substation South _Attic` (22 chars) matched `Substation South` (16) at
distance 6 and the suggestion **silently dropped "Attic"**. A correction
is written straight into the sheet, so it has to be a typo-level
difference and nothing more.

Catches: `Cooridor` → `Corridor`, `Elctrical Room` → `Electrical Room`,
`Garabge room` → `Garbage Room`.

---

## 6. Path handling — `core/paths.py`

### 6.1 The three faults

```
"Z:\WIR request\..."    wrapped in quotes (Explorer's "Copy as path")
S:\WIR request\...      the share mapped to a different drive letter
Z:\WIR request\...      the \Common\ level missing
```

### 6.2 `repair`

1. Strip quotes — straight, curly, apostrophes.
2. Normalise spaces, strip trailing backslashes.
3. If the drive letter is **C–H**, return untouched. Those are real local
   folders, not a mis-mapped share.
4. UNC paths (`\\server\...`) are returned untouched.
5. Otherwise force the drive to `Z:` and ensure the path sits under
   `Z:\Common\`.

### 6.3 `repair_cell` — the conservative rule

For each path in the cell:

- **Resolves as written** → left alone. If only the quotes were wrong,
  the unquoted form is written back.
- **Does not resolve** → try the repair. Written back **only if the
  repaired path actually exists**.
- **Neither resolves** → the original text is left *exactly* as typed and
  the row is flagged. At that point it needs a person, not a guess.

Multiple paths in one cell keep their original separators, because each
path is replaced in place inside the text rather than the cell being
rebuilt.

### 6.4 `share_available` — threaded with a timeout

```python
_MAX_WAIT = 3.0 seconds, cached for the process
```

A single `os.path.isdir()` against a disconnected mapped drive blocks for
about **twenty seconds** while Windows retries the connection. That was
the entire runtime of a check when the share was down.

### 6.5 `is_network_path`

A UNC path, or any drive letter that is not local. Checked before probing
when the share is down — a link typed as `W:\Common\...` is just as dead
as one on `Z:`, and each disconnected probe costs twenty seconds.

### 6.6 Why this matters

When the share is unreachable, **every** path under it fails to resolve.
Reporting 52 broken links then is worse than useless: it buries the real
ones and flags rows that are fine. Both Fix Attachment Links and
pre-flight ask once and, if the answer is no, say so instead of blaming
the data.

---

## 7. The tasks

### 7.1 Generate WIRs

Output: the `WIRs` folder beside the workbook. Not configurable from the
UI — it is always the same place, and being asked every time is a step to
get wrong.

Per row, skipping blanks and anything marked `Done`:

1. **Folder** — `<output>\<WIR No>-R<rev>\`, rev from column O padded to
   two digits.
2. **Cover** — write the WIR number into `WIR-Form!BB1`, recalculate,
   export to `<WIR No>-R<rev>.pdf`. That one cell drives every field on
   the form.
3. **Checklist** — sheet named in column P, same `BB1` mechanism, export
   to `<WIR No>-XCKL.pdf`. `Comments_Form` gets its print area
   recalculated first (`A1:J<last>`) because it grows with its content.
**Import Previous WIR Paths runs first, always.** Pressing Generate WIRs
calls `tasks.generate_wirs`, which refreshes column AA and only then
builds. Column AA is right only as of the last register import, and a WIR
that goes out without its predecessor covers comes back rejected — so the
step is not left to memory. Its own button stays, for filling the column
without generating. If the refresh fails, Generate does not run:
generating against paths known to be stale is worse than not generating.

4. **Attachments** — columns W, U, Y, AA in that order, becoming
   `-XTTCH1`…`-XTTCH4`. Each cell may hold several paths; each path may be
   a file or a folder, and a folder contributes all its files with the
   numbering continuing across the cell. Broken paths get the same repair
   Fix Attachment Links uses before being given up on.

   | Column | Header | Constant | Copied |
   |---|---|---|---|
   | W | Attachment QC | `ATT_QC` | whole |
   | U | Attachment Site | `ATT_SITE` | whole |
   | Y | Previous WIR Rev. | `ATT_PREV` | whole |
   | AA | Previous Activity | `ATT_FIRST` | **first page only** |

   The last two constant names read backwards and have caused a real bug
   — `ATT_PREV` is the earlier *revision* of this same WIR (attached when
   column O is above 0), while `ATT_FIRST` is column AA, the approved
   *predecessor activities* that Import Previous WIR Paths fills. Check
   the column letter, never the name.

   Column AA takes **only the cover page** of each predecessor: these are
   whole commented WIRs, often 30–45 pages each. On a real row, four
   predecessors totalling 147 pages contributed 4. Output naming is
   unaffected — still `<WIR No>-XTTCH4-<n>.pdf`.

   A PDF that cannot be read falls back to being copied whole. A WIR
   carrying too many pages is a nuisance; one missing its evidence is a
   rejection.
5. **Merge** — every PDF in the row's folder into one
   `<WIR No>-R<rev>.pdf` in the output root.
6. **Mark the row `Done`.**

Merge order is by filename, which happens to be correct:
`-R00` → `-XCKL` → `-XTTCH…`, since `R` sorts before `X`.

**Junk files are skipped**: `Thumbs.db`, `desktop.ini`, `.DS_Store`,
`ehthumbs.db`, and anything with the hidden or system attribute. These
were being copied into every WIR *and* blocking the folder cleanup.

### 7.2 The merge contract — `core/pdf.py`

| Outcome | Merged file | Row folder |
|---|---|---|
| Everything was a PDF | created | **deleted** |
| A non-PDF was present | created | **kept** — it could not be merged |
| Merge failed | **removed** | **kept**, untouched |
| No PDFs at all | not created | kept |

On failure the partial output is deleted. A half-written PDF sitting in
the output folder looks like a finished document, and someone will submit
it.

The original `wir_generate.py` deleted the folder unconditionally after
merging only `*.pdf`, which silently destroyed photos and any other
non-PDF attachment.

### 7.3 Fix Attachment Links

Column N of the log only. U, W, Y and AA are formulas derived from it.

Rows marked `Done` are skipped. The column's fill is cleared first, or a
link fixed since the last run stays red. If the share is unreachable the
whole pass is skipped and nothing is coloured.

Rows where a path still cannot be resolved are filled red and listed.

### 7.4 Fix WIR Path Prefix

Column N of the register. Pure text substitution, no disk access:

```
\\192.168.225.6\i125MEP\QC\1. Work Inspection Request\Commented WIR\ELE\
  ->  D:\i125\Logs\WIR\ELE&ELV\extracted\
```

**Runs automatically at the end of Import WIR Data**, which is the only
thing that puts those paths there. Also available on its own button for
when the register is edited by hand.

### 7.5 Check Areas

Writes into **column K and nowhere else**. The Area column is never
overwritten, so nothing is lost if a suggestion is wrong.

Per row, skipping `Done` and any floor not in the checkable set:

- **Suggestion** — the Area rewritten into house format. Blank if it
  already matches.
- **Note** — what could not be resolved.
- Both, on separate lines, when both apply.
- The cell is filled amber when there is a note.

Checkable floors: `BS, GF, 1F–10F, RF`. Everything else — `SC01`, `SC02`,
`RS`, `PD`, `UR`, and typos like `1F-2F` — is skipped entirely, because
there is no agreed area scheme to check against.

Formatting rules:

```
301 302 303 304                  -> Apartments 301, 302, 303, 304
101,102,103  (stored as 101102103) -> Apartments 101, 102, 103
Apartments 401  402,403          -> Apartments 401, 402, 403
Apartment no 401, 402            -> Apartments 401, 402
Flats 301, 302                   -> Apartments 301, 302
Cooridor                         -> Corridor
Garbage Room and Water Meter Room -> Garbage Room, Water Meter Room
Telephone-Water Meter Room       -> Telephone Room, Water Meter Room
Telephone room Garbage room      -> Telephone Room, Garbage Room
Part 1,5 - Building C1 (Package 08) -> Part 1, Part 5
Mock up Electrical Room          -> Mock up Electrical Room
503 on 4F                        -> unchanged, out of band
```

Six rules beyond the obvious, each from data that broke the previous
version:

1. **Unit words repeat.** `Apartment no 401` is two of them in a row.
   `match_key` strips them in a loop — stripping only the first leaves
   `NO 401`, which is not a unit code, so the number was never checked.
2. **Hyphens separate, digits protect.** `Telephone-Water Meter Room` is
   two rooms; `Part-5` and `A1-004` are one thing each. The split runs
   only where no digit sits on either side of the hyphen.
3. **Package and Building are stripped** from the Area — they have their
   own columns, and typed here they pushed the real content out of reach.
4. **Two names with no separator are split** (`split_known`), but only
   when *every* part resolves. A partial match would invent an area
   nobody wrote. The tail is re-split, so `water meter room corridor LL`
   yields both rooms.
5. **Mock up is a prefix, not an area.** It is stripped, the remainder is
   resolved, and it is put back — `Mock up Electrical Room` is the
   electrical room. `Mock-up` is normalised *before* the hyphen split, or
   it becomes `Mock` and `up`.
6. **Staircases are corrected in the Floor column, not the Area.**
   `Staircase 1`, `staircase 01`, `Stair case 2`, `STAIRCASE-1` all carry
   a number, and it belongs in Floor as `SC01` / `SC02`. A staircase with
   no number is reported as such.
7. **A unit number with a room name stays whole.** `G10 Kitchens` is a
   unit and a room in one token. Splitting it would have to guess which
   half is the area, so it is kept exactly as typed and reported —
   free text with a note, not a correction.
8. **`.R` expands to `Room`** — `Garbage .R` is Garbage Room. Expanded
   rather than added as a variant, because the abbreviation is a habit
   and not specific to one room.
8. **Free text gets no suggestion at all.** `format_area` leaves an
   ignorable token exactly as typed. Suggesting something here is worse
   than suggesting nothing: `GF to First Floor columns` came back as
   `Gym, to First Floor columns`, and a wrong suggestion is one someone
   may copy into the Area.

`nearest` will not fuzzy-match a probe shorter than
`_MIN_FUZZY_LENGTH` (4). Two edits is most of a three-letter word, so
short tokens matched almost anything — `GF` became `Gym`. The
abbreviations that genuinely mean a room (MTR, BMS, CBS, GSM, RMU) are
listed in `Area_Names` and found by exact lookup, so they never used the
fuzzy path.

Three vocabularies in `config` decide what is accepted without comment:

| | |
|---|---|
| `SYSTEM_WORDS` | CCTV, Lighting, Power … — systems, not places. `BMS` is deliberately absent; BMS Room is a real room. |
| `FREE_TEXT_MARKERS` | HIGHLIGHTED, SLAB, RAFT, COLUMN … — the value describes the extent of work. Floor ranges (`2nd to 3rd`) match too, via `is_floor_range`. |
| `IGNORED_AREAS` | `Lift Core`, `Service rooms` — real, deliberately not checked. Matched against `match_key`, which has already dropped a trailing `ROOM`/`ROOMS`, so `Service rooms` arrives as `SERVICE` while `Service corridor` stays a real area. |

A leftover shorter than `MIN_REPORTABLE_AREA` (4) is not reported — it is
a fragment from a split, not a name anyone typed.

Unit codes are collected under one `Apartments` prefix **in the order they
were typed** — not sorted, because the typed order sometimes carries
meaning and reordering it silently would be surprising.

A number is only called an apartment when the text says so
(`has_unit_prefix`) or it matches the floor's band (`is_unit_for_floor`)
or it is a recovered glued run. Without that, `Staircase 1, 2` became
`Apartments 2`.

Notes carry no boilerplate. When a close match exists it is applied in
the suggestion rather than asked about; only what cannot be corrected
gets a note.

### 7.6 Check WIRs Before Generate (pre-flight)

Every validation, before a single file is written.

| Check | Severity | Fires when |
|---|---|---|
| WIR number | ERROR | length ≠ 42, or package/dep. inside the number disagree with columns F/E |
| Duplicate number | ERROR | same number twice in the log |
| Already submitted | ERROR | this number is already in the register |
| Missing fields | ERROR | Dep./Package/Building/Floor/Activity empty |
| Rev | ERROR | column O is not a number |
| Checklist | ERROR | column P empty, or names a sheet that is not in the workbook |
| Description | WARN | column C empty |
| Previous activity | WARN | a required predecessor has no approved WIR covering every area |
| Area submitted before | WARN | same activity already approved here for one of this row's areas |
| Apartment | WARN | a unit is not in the register for this floor |
| Area name | WARN | not in `Area_Names`, or on a floor it cannot be on |
| Attachment | WARN | a given path does not resolve |

**Already submitted** matches on the WIR number alone. The location-based
check deliberately ignores an entry carrying the row's own number so a row
does not warn about itself after a register re-import — which left the
strongest duplicate signal of all with nothing to catch it.

```
same rev, approved     -> ERROR "would duplicate an approved WIR"
same rev, other status -> ERROR "raise the revision before resubmitting"
higher rev, approved   -> WARN  "resubmitting work already signed off"
```

**Previous activity** is the most involved check. An activity counts as
done for one area if:

1. it has an approved WIR of its own, **or**
2. it is a sub-activity and the parent was raised covering it, **or**
3. it is a parent and **every** one of its subs is approved.

Both directions are needed: sometimes the main activity is raised to cover
both subs, sometimes the subs are raised to cover the main.

Uncovered areas are named and grouped by *why*:

```
At P09-C-009-2F
Wires_Pulling (raised as Wires_H.L + Wire_PWR_Pulling) [Not submitted (201, 202)]
DB&ONU_Enclosure (raised as DB_Enclosure) [Rejected (206, 207)]
```

Status comes from the register: an existing but unapproved WIR reports its
status word; nothing at all reports `Not submitted`.

Where the table lists a parent *and* its subs as separate predecessors and
none are approved, only the parent is named — otherwise the same fact is
stated three times.

**Attachment coverage** is a count, not a per-row warning, for the QC
column: it points at a folder that only exists once QC has filed
something, so a missing one is normal rather than a fault.

### 7.7 The PreFlight sheet

**One line per log row, not per finding.** A row with a duplicate number
*and* a precedence gap is an ERROR row: the worst finding sets the
severity, and the comment carries every finding with the errors first, so
the blocking problem is never buried under a warning.

```
Row | WIR No. | Dep. | Package | Building | Floor | Area | Activity | Severity | Check | What to fix
```

Error rows sort above warning rows, each block still in log order. The row
number links back to the log sheet. Panes are not frozen.

On the log sheet itself, **only columns B and I are coloured**, and both
are cleared first — so this can never disturb the attachment colouring in
U/W/Y/AA. Red for errors, amber for warnings. Column I is coloured only
for findings about the row's own area.

### 7.8 Import WIR Data

Opens the master register read-only, clears `WIRs!A:AD`, and stacks rows
`A:T` of the `ELE` and `ELV` sheets from row 4. Filters are cleared on the
source first.

The four derived columns are **computed in Python**, not left as
formulas — the workbook has no macros now, and `AD` used to depend on one.

Runs Fix WIR Path Prefix at the end.

### 7.9 Import Previous WIR Paths

For each visible, pending row: find the approved WIRs for the activities
that must precede this one at the same Package-Building-Floor, and write
their paths into column AA (**Previous Activity**), semicolon-separated.

Paths come from the register's **column N** — the path as filed, on the
share. Changed 2026-08-17 from column AD, the local extract folder. The
consequence is that this task and Generate both need the server
connected; AD did not. Generate takes only the cover page of these files,
so the size of the originals no longer matters.

Split-aware in both directions. Area-aware: a register entry counts if its
area overlaps the row's, if it is blank (covered everything), or if it
says "All Apartments" and the row names a unit.

The column is **seeded with what is already there**. A row this pass skips
— hidden, marked Done, missing the fields to match on, or with no
predecessors defined — keeps its existing paths. Only a row that was
actually examined is rewritten.

### 7.10 Copy To History

Copies **visible rows only** — a filtered log is how you choose what goes
out, so hidden rows are not history. Appends columns A:AC to
`WIR_History` with a timestamp in AD, creating the sheet and its header
row if needed.

`WIR_History` is laid out **exactly like the log**: row 1 blank, headers
on row 2 (`config.HEADER_ROW`), data from row 3. This matters more than it
looks. An earlier version read the header from row 1, found it empty on a
sheet holding 660 rows, concluded the sheet was new, and would have
appended from row 2 — over the header and the entire history. The append
row is now `max(last_row + 1, FIRST_DATA_ROW)`, so a surprise in the sheet
can no longer cost data.

Dates are re-formatted `dd-mmm-yyyy` after writing, because rows are read
with `Value2` and a date arrives as an Excel serial number.

The timestamp is written as an **Excel serial number**, not a string, and
the column is formatted `yyyy-mm-dd hh:mm:ss`. The existing rows are
stored that way, and a date column with text mixed in cannot be sorted or
filtered. Computing the serial directly from `EXCEL_EPOCH`
(1899-12-30) also keeps this off the COM date-conversion path that caused
the original `win32timezone` failure.

Optionally opens an Outlook draft with a table of columns B–J plus O. It
**displays** the mail, never sends it, and leaves the recipient empty. All
cell values are HTML-escaped.

The window asks at the moment the button is pressed — Yes drafts the mail,
No copies to history only, Cancel abandons both. This replaced a checkbox
on the Actions tab, which sat too far from the button to be noticed.

`_open_email` returns `True`, or the reason it failed. A reason is a
non-empty string, so the caller must test `sent is True` — testing it for
truth reports every Outlook failure as a success.

`greeting` is supported and rendered above the table, but the window
passes `""`; there is no field for it, on the assumption that the draft is
edited in Outlook anyway.

### 7.11 Extract PDFs

Copies whatever the visible rows' Link column points at into one chosen
folder. Original filenames are kept, numbered `name (1).ext` on collision
rather than overwriting.

The VBA asked the user to select a range by hand; this uses the same
filtered set Copy To History works on.

### 7.12 Clear sheets

**Clear Any Sheet** — contents and fills of rows 3–1000, columns A:AZ, on
whatever sheet is active. Borders, fonts, number formats and column widths
are untouched.

**Clear WIR Sheet** — the same, then restores the template row down to row
100: the number-check formula, description, apartments, location, form,
the three hyperlink columns, today/tomorrow dates, and Rev 0.

Both ask for confirmation in the UI.

> **Consequence of removing the VBA**: columns V, X, Z and AB previously
> held `=PathExists(...)`, a function that lived in the workbook's
> `Functions` module. With no macros those show `#NAME?`, so they are left
> empty. Fix Attachment Links and Check WIRs both report which paths
> resolve, which is what those columns were for.

---

## 8. The window

PySide6. Two tabs: **Actions** and **Last run**.

Three colour-coded sections:

| Section | Colour | Buttons |
|---|---|---|
| Tools | teal | Import WIR Data, Import Previous WIR Paths, Extract PDFs, Clear Any Sheet |
| Verify and maintain | purple | Fix Attachment Links, Fix WIR Path Prefix, Check Areas, Check WIRs Before Generate |
| Generate | coral | Generate WIRs, Copy To History, Clear WIR Sheet |

Plus **Run all checks** — Fix Attachment Links → Check Areas → Check WIRs,
in that order, because each feeds the next: repaired links stop the
attachment check reporting phantom failures, and the area pass puts its
suggestions in K before pre-flight reads the rows.

Fix WIR Path Prefix is **not** in Run all checks — it belongs to the
register, and Import WIR Data already runs it.

### 8.1 Threading

Every task runs on a `QThread` worker so the window never freezes. COM is
initialised inside the worker (`pythoncom.CoInitialize`) and every COM
object is created and used on that same thread.

**Dialogs are opened on the GUI thread** and the result handed to the
worker. A file dialog from a worker thread is a crash waiting to happen.

### 8.2 Destructive actions

`Clear Any Sheet` and `Clear WIR Sheet` ask for confirmation first.

---

## 9. Performance

COM crosses a process boundary. The discipline that matters:

- **Read a whole range in one call, write it back in one call.** Reading
  the 2,193-row register takes 0.04s as a block.
- **Cell-by-cell loops from Python are slow** in a way they never were in
  VBA.
- `sheet_exists()` answers from a cached set of names — pre-flight asks it
  once per row for the checklist in column P, and walking `Worksheets`
  each time is 60 × 30 COM calls for an answer that cannot change.

Measured on the live workbook (61 log rows, 2,193 register rows):

```
attach                 0.05s
reference.load         0.19s
read log block         0.01s
read register block    0.04s
Fix Attachment Links   0.21s
Check Areas            1.48s
Pre-flight             4.25s
full --check           3.4s   (share connected)
Generate, 3 rows      ~26s    (Excel rendering dominates)
```

### 9.1 `Value2`, never `Value`

`Range.Value` converts a date cell into a COM date, which makes pywin32
import `win32timezone` at that moment — a module that is missing unless
pywin32's post-install step has run, and one PyInstaller cannot see.

The log has dates in columns S and T, so **every** task that read a row
block hit it. `Value2` returns dates as plain Excel serial numbers, so the
dependency never arises. It is faster too.

`text_of()` exists because Excel hands back floats: `101102103.0` must not
become the string `"101102103.0"`.

---

## 10. Shipping it

Two routes. **The portable folder is the one to use.** The `.exe` is kept
as a fallback and is documented after it.

There are deliberately no `.bat` or `.cmd` files anywhere in this project.
Batch files draw security attention in this environment, and an unsigned
PyInstaller `.exe` draws more — antivirus quarantines them routinely. The
portable folder avoids both: the only executable in it is Python's own
`python.exe`, signed by the Python Software Foundation.

### 10a. The portable folder (preferred)

```
python tools\deploy.py --final-path "C:\WIR Generate Tools"
```

Produces a folder that needs no installation, no compilation and no admin
rights:

```
WIR Generate Tools\
    WIR Generate Tools.lnk         double-click to open the window
    Check this PC.lnk     runs --selftest, for a machine's first use
    runtime\              a copy of the embeddable Python (~250 MB)
    app\                  main.py, core, ui - readable source
    README.txt
    WIR Generate Tools - User Manual.docx
```

**A `.lnk` can only hold an absolute path.** There is no portable form of
one. Windows will sometimes re-find a target that has moved, but that is
a heuristic, not a guarantee, and it is not something to rely on for
other people's machines.

So the folder needs one agreed location, passed as `--final-path`, and
the shortcuts are written for it. Either:

- **a folder on the share**, if the drive letter is the same for
  everyone — best option, because it is also one copy to update rather
  than one per machine; or
- **the same local path on every PC**, e.g. `C:\WIR Generate Tools`.

Build in place by omitting `--final-path`.

Two details make the shortcuts sturdier than they look. The working
directory is left empty, so Explorer starts the process in
`pythonw.exe`'s own folder; the script is then named `..\app\main.py`
relative to that. Only the runtime path has to be right, not two paths.
And `WIR Generate Tools.lnk` targets `pythonw.exe` rather than `python.exe`, so no
console flashes behind the window.

Users should copy the *shortcut* to their desktop, not move the folder.

### 10b. The executable (fallback)

Needs a machine with Python. Two commands, no batch file:

```
pyinstaller --noconfirm --clean "WIR Generate Tools.spec"
"dist\WIR Generate Tools.exe" --selftest
```

The second is not optional — it is what turns "it built" into "it will
run over there". Delete `build\` and `dist\` first if a previous build
exists, or PyInstaller reuses stale analysis.

### Why a spec file

PyInstaller decides what to bundle by reading the source for `import`
statements. That misses anything imported by name at runtime, which is
precisely how pywin32 behaves:

- `pythoncom` imports **`win32timezone`** the first time COM hands back a
  date. Nothing in our source ever mentions it, so PyInstaller has no way
  to know — and the failure surfaces on the user's machine, mid-run.
- `win32com.client.Dispatch` builds its wrappers from strings.

`WIR Generate Tools.spec` lists these by hand and adds
`collect_submodules("win32com")`, `("win32comext")`, `("pypdf")` and
`("cryptography")`. It deliberately over-includes: a few extra megabytes
cost nothing, a missing module costs a site visit.

`cryptography` is there because `pypdf` only imports it when it meets a
password-protected PDF, which consultants' attachments often are.

Excludes are kept to `tkinter`, `unittest`, `test`, `lib2to3`,
`pydoc_data`. Trimming Qt modules to save size is not worth it — every
exclusion is a chance to break the build in a way that only appears on
someone else's machine.

`upx=False`: UPX compression corrupts some Qt DLLs.

### The self-test

`--selftest` needs no workbook and no network. It runs five checks:

| Check | What it proves |
|---|---|
| program modules | all 21 of our own modules made it into the bundle |
| third-party modules | pywin32, pypdf and PySide6 are present |
| COM dates | `Scripting.FileSystemObject.DateLastModified` — the exact path that used to raise `No module named 'win32timezone'` |
| Qt platform plugin | Qt starts and finds `qwindows.dll`, which is a DLL rather than an import and so is missed by import checks |
| PDF merge | builds two PDFs and merges them, as Generate does |

The last three do real work rather than importing, because that is where
frozen builds fail. The report is printed and written to `selftest.log`
beside the executable — the release build is `console=False`, so printing
alone would go nowhere.

Run it on a target machine after copying, and it confirms that machine can
run the program before anyone tries a real batch. In the portable folder
that is what **Check this PC** does; `--pause` keeps the console open long
enough to read, and the report is written to `app\selftest.log`.

### If startup fails anyway

`main.py` wraps `main()` and shows a message box naming the failure and
telling the user to run `--selftest`. It uses `ctypes.MessageBoxW` rather
than Qt, because the most likely reason for getting there is a missing
piece of the bundle and Qt is a candidate for being that piece.

### Notes

- Drop `console=False` in the spec while debugging, or a crash in a worker
  thread vanishes silently.
- `--check` gives the headless output from a terminal, either as
  `WIR Generate Tools.exe --check` or `runtime\python.exe app\main.py --check`.
- The target machine needs the Microsoft Visual C++ 2015–2022
  redistributable. PyInstaller bundles `vcruntime140.dll`, and any machine
  running current Office already has it, so this has not been a problem in
  practice.
- `WIR Generate Tools.spec` is checked in — do not let PyInstaller regenerate it.
- Do not commit `build/`, `dist/` or `selftest.log`.

---

## 11. Known limitations

1. **Hyphens are not separators.** `Service rooms Electrical - Telecom -
   WMR - Service corridor` comes through as one unknown token. Retyping
   with commas is the fix; splitting on hyphens would break `Part-5` and
   `Mock-up`.
2. **Compound areas are not decomposed.** `Substation South _Attic` is one
   token. It is reported, not corrected, because auto-correcting it would
   have to drop half the value.
3. **Mixed tokens stay whole.** `G10 Kitchens` is not split into a unit
   plus a room.
4. **Excel destroys bare comma lists** at typing time. `101,102,103`
   becomes the number `101102103`. Recovery works only when every 3-digit
   chunk lands in the floor's band. Typing `Apartments` first keeps the
   cell as text.
5. **Areas are validated against a flat vocabulary**, with floor rules
   only where a Floors value is set. There is no per-building-floor room
   list.
6. **Only ELE and ELV are imported** from the master register.
   `Prev_Activities` and `Activity_Splits` cover those trades only, so PUB
   rows get no precedence checking — the check is silently skipped rather
   than warning.
7. **The register must be re-imported** for any check to see recent
   approvals.

---

## 12. Decision log

Things that look odd and are deliberate.

| Decision | Reason |
|---|---|
| COM instead of openpyxl | must edit the workbook while it is open |
| `GetActiveObject` not `Dispatch` | Dispatch starts a second empty Excel |
| `Value2` everywhere | avoids the `win32timezone` import entirely |
| Log sheet cached per connection | a mid-run tab change must not redirect writes |
| Reference tables by name | the data has moved sheets twice already |
| Check Areas writes to K, not I | a wrong suggestion must not destroy the Area |
| Fuzzy match capped at 2 edits + 2 length | a looser cap silently truncated a value |
| Hyphens not separators | 30 of 1,788 areas rely on them |
| Portable folder, not an .exe | an unsigned PyInstaller build gets quarantined; `python.exe` is signed |
| No `.bat` or `.cmd` anywhere | batch files draw security attention in this environment |
| Shortcuts pinned to one agreed path | a `.lnk` cannot hold a relative target; relocation is a heuristic, not a guarantee |
| Share probed once, on a thread | a dead mapped drive costs 20s per probe |
| Local drives C–H never rewritten | they are real folders, not a mis-mapped share |
| Repair written back only if it resolves | an unverified guess is worse than a flag |
| Folder kept when merge fails | manual merge needs the parts |
| Partial merge output deleted | it looks like a finished document |
| Junk files skipped | they blocked folder cleanup and polluted output |
| One report line per row | the blocking error must not hide under a warning |
| Main/Essential/Electrical kept separate | they are different rooms on this project |
| `MTR` = Main Telephone Room | all 8 register rows are in the basement |
| Import Previous seeds from existing | it must not wipe rows it chose to skip |
| Generate output not configurable | it is always the same folder |

---

## 13. Verified against live data

| Area | Status |
|---|---|
| COM attach, workbook lookup | verified |
| All five reference tables by name | verified — 33 activities, 2 splits, 123 floor codes, 6 status codes, 46 areas |
| `Value2` reads | verified |
| Fix Attachment Links | verified, including a real `W:` → `Z:` repair |
| Share unreachable handling | verified in both states |
| Check Areas | verified — 12 formatter cases, 11 live findings |
| Pre-flight | verified — 61 rows, report sheet, colouring |
| Generate | verified — 3 rows: covers, checklists, `Comments_Form`, attachments, merge, Done |
| Merge contract | verified — all four outcomes |
| Import Previous WIR Paths | verified — 11 rows, idempotent |
| Fix WIR Path Prefix | verified standalone |
| Import WIR Data | **not yet run** |
| Copy To History, Outlook draft | **not yet run** |
| Extract PDFs | **not yet run** |
| Clear sheets | **not yet run** |
| The Qt window | **not yet run** |
