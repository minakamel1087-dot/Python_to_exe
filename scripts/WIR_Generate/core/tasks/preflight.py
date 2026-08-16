"""
Pre-flight - every check, before a single file is written.

Generating finds problems the expensive way: it writes the PDFs, colours
a cell, and moves on. By then the documents exist, the row is marked Done
and the mistake is already in the submission folder. This asks the same
questions first.

Output: the PreFlight sheet, one line per log row with the errors ranked
first, plus red/amber on the WIR number and Area cells.
"""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field

from .. import config
from ..findings import AREA_CHECKS, Finding, RunResult, Severity
from ..paths import exists, repair, split_paths, trim_slashes, unquote
from ..reference import ReferenceData
from ..text import is_unit_code, match_key, nearest, normalize_spaces, tokenize
from ..workbook import LiveWorkbook, text_of
from .check_areas import floor_is_checkable

TITLE = "Check WIRs Before Generate"

NOT_SUBMITTED = "Not submitted"

ATTACHMENT_COLUMNS = [
    # column, label, warn when a given path does not resolve
    (config.Main.ATT_QC, "QC Attachment (W)", False),
    (config.Main.ATT_SITE, "Site Attachments (U)", True),
    (config.Main.ATT_PREV, "Previous REV (Y)", True),
    (config.Main.ATT_FIRST, "Previous WIR paths (AA)", True),
]


@dataclass
class _RegisterEntry:
    wir_no: str
    area: str
    status: str
    rev: int


@dataclass
class _Index:
    """The register, arranged for the questions this file asks."""
    by_location: dict[str, list[_RegisterEntry]] = field(default_factory=lambda: defaultdict(list))
    by_number: dict[str, _RegisterEntry] = field(default_factory=dict)


def _key(package: str, building: str, floor: str, activity: str) -> str:
    return f"{package}|{building}|{floor}|{activity}".upper()


def _build_index(rows: list[list], ref: ReferenceData) -> _Index:
    index = _Index()
    for row in rows:
        if len(row) < config.Register.LAST_COL:
            row = row + [None] * (config.Register.LAST_COL - len(row))

        wir_no = text_of(row[config.Register.WIR_NO - 1])
        status = text_of(row[config.Register.STATUS - 1])
        rev_text = text_of(row[config.Register.REV - 1])
        rev = int(float(rev_text)) if rev_text.replace(".", "").isdigit() else 0

        entry = _RegisterEntry(
            wir_no=wir_no,
            area=text_of(row[config.Register.AREA - 1]),
            status=status,
            rev=rev,
        )

        package = text_of(row[config.Register.PACKAGE - 1])
        building = text_of(row[config.Register.BUILDING - 1])
        floor = text_of(row[config.Register.FLOOR - 1])
        activity = text_of(row[config.Register.ACTIVITY - 1])
        if package and building and floor and activity:
            index.by_location[_key(package, building, floor, activity)].append(entry)

        if wir_no:
            # Where a number appears more than once, the highest revision
            # wins - that is the state the number is actually in.
            seen = index.by_number.get(wir_no)
            if seen is None or rev >= seen.rev:
                index.by_number[wir_no] = entry

    return index


# --- area matching ---------------------------------------------------------


def _area_covers(entry_area: str, token: str) -> bool:
    """Does an approved WIR's area cover this one area?

    A blank area on the register row means that WIR covered everything;
    "All Apartments" covers any unit number.
    """
    if not token:
        return True
    if not entry_area.strip():
        return True
    if "ALL APARTMENT" in entry_area.upper():
        return is_unit_code(token)
    return token in {match_key(t) for t in tokenize(entry_area)}


def _covering(index: _Index, package: str, building: str, floor: str,
              activity: str, token: str, ref: ReferenceData,
              approved_only: bool = True, skip_number: str = "") -> _RegisterEntry | None:
    for entry in index.by_location.get(_key(package, building, floor, activity), []):
        if skip_number and entry.wir_no.upper() == skip_number.upper():
            continue
        if approved_only and not ref.is_approved(entry.status):
            continue
        if not approved_only and ref.is_approved(entry.status):
            continue
        if _area_covers(entry.area, token):
            return entry
    return None


def _covered(index: _Index, package: str, building: str, floor: str,
             activity: str, token: str, ref: ReferenceData) -> bool:
    """An activity counts as done for one area if any of these hold:
      - it has an approved WIR of its own
      - it is a sub-activity and the parent was raised covering it
      - it is a parent that gets raised as sub-activities, and every one
        of those subs is approved (no WIR for the parent will ever exist)
    """
    if _covering(index, package, building, floor, activity, token, ref):
        return True

    for parent in ref.split_parents.get(activity, set()):
        if _covering(index, package, building, floor, parent, token, ref):
            return True

    subs = ref.split_subs.get(activity)
    if subs:
        return all(
            _covering(index, package, building, floor, sub, token, ref) for sub in subs
        )

    return False


def _token_status(index: _Index, package: str, building: str, floor: str,
                  activity: str, token: str, ref: ReferenceData) -> str:
    """What the register says about this activity/area when it is not
    approved: the status of a WIR that exists but is not signed off, or
    "Not submitted" when there is no WIR for it at all."""
    candidates = [activity]
    candidates.extend(sorted(ref.split_subs.get(activity, set())))
    candidates.extend(sorted(ref.split_parents.get(activity, set())))

    for candidate in candidates:
        entry = _covering(index, package, building, floor, candidate, token, ref,
                          approved_only=False)
        if entry is not None:
            return ref.status_text(entry.status)

    return NOT_SUBMITTED


# --- the run ---------------------------------------------------------------


class _Run:
    def __init__(self, book: LiveWorkbook, ref: ReferenceData):
        self.book = book
        self.ref = ref
        self.findings: list[Finding] = []
        self.row_severity: dict[int, Severity] = {}
        self.area_severity: dict[int, Severity] = {}
        self.row_context: dict[int, tuple] = {}
        self.attachments = {label: [0, 0] for _, label, _ in ATTACHMENT_COLUMNS}

    def add(self, row: int, wir_no: str, severity: Severity, check: str, message: str) -> None:
        self.findings.append(Finding(row, wir_no, severity, check, message))

        current = self.row_severity.get(row)
        if current is None or severity > current:
            self.row_severity[row] = severity

        if check in AREA_CHECKS:
            current = self.area_severity.get(row)
            if current is None or severity > current:
                self.area_severity[row] = severity

    # -- individual checks --------------------------------------------------

    def check_number(self, row: int, wir_no: str, package: str, dep: str,
                     seen: dict[str, int]) -> None:
        if len(wir_no) != config.WIR_NO_LENGTH:
            self.add(row, wir_no, Severity.ERROR, "WIR number",
                     f"Length is {len(wir_no)}, expected {config.WIR_NO_LENGTH}.")
        else:
            if wir_no[17:20].upper() != package.upper():
                self.add(row, wir_no, Severity.ERROR, "WIR number",
                         f"Package in the number is '{wir_no[17:20]}' "
                         f"but column F says '{package}'.")
            if wir_no[28:31].upper() != dep.upper():
                self.add(row, wir_no, Severity.ERROR, "WIR number",
                         f"Dep. in the number is '{wir_no[28:31]}' "
                         f"but column E says '{dep}'.")

        if wir_no in seen:
            self.add(row, wir_no, Severity.ERROR, "Duplicate number",
                     f"This WIR number is already used on row {seen[wir_no]}.")
        else:
            seen[wir_no] = row

    def check_already_submitted(self, row: int, wir_no: str, rev: int, index: _Index) -> None:
        """This exact number is already in the register.

        The location check below ignores an entry carrying the row's own
        number, so that a row does not warn about itself after a re-import.
        That left the strongest duplicate signal of all - the same number
        going out twice - with nothing to catch it.
        """
        entry = index.by_number.get(wir_no)
        if entry is None:
            return

        state = f"Rev {entry.rev}, {self.ref.status_text(entry.status)}"
        approved = self.ref.is_approved(entry.status)

        if rev > entry.rev:
            if approved:
                self.add(row, wir_no, Severity.WARN, "Already submitted",
                         f"The register already has this number approved at {state}. "
                         f"Rev {rev} would resubmit work that is already signed off.")
            return

        if approved:
            self.add(row, wir_no, Severity.ERROR, "Already submitted",
                     f"This number is already in the register as {state}. "
                     f"Generating it again at Rev {rev} would duplicate an approved WIR.")
        else:
            self.add(row, wir_no, Severity.ERROR, "Already submitted",
                     f"This number is already in the register as {state}. "
                     "Raise the revision before resubmitting it.")

    def check_precedence(self, row: int, wir_no: str, package: str, building: str,
                         floor: str, area: str, activity: str, index: _Index) -> None:
        required = self.ref.previous.get(activity)
        if not required:
            return

        tokens = [t for t in tokenize(area)] or [""]
        missing: dict[str, str] = {}

        for previous in sorted(required):
            gaps: dict[str, list[str]] = defaultdict(list)
            covered_any = False

            for token in tokens:
                key = match_key(token)
                if _covered(index, package, building, floor, previous, key, self.ref):
                    covered_any = True
                    continue
                status = _token_status(index, package, building, floor, previous,
                                       key, self.ref)
                gaps[status].append(token or "")

            if not gaps:
                continue

            if not covered_any and all(not t for t in tokens):
                missing[previous] = next(iter(gaps))
            else:
                parts = []
                for status, areas in gaps.items():
                    named = ", ".join(a for a in areas if a)
                    parts.append(f"{status} ({named})" if named else status)
                missing[previous] = "; ".join(parts)

        if not missing:
            return

        # The table often lists a parent AND its sub-activities separately.
        # When none of them are approved, naming all three says the same
        # thing three times - keep the parent and note what it stands for.
        redundant = {
            name for name in missing
            if any(parent in missing for parent in self.ref.split_parents.get(name, set()))
        }

        lines = [f"At {package}-{building}-{floor}"]
        for name, gap in missing.items():
            if name in redundant:
                continue
            subs = [s for s in self.ref.split_subs.get(name, set()) if s in missing]
            label = f"{name} (raised as {' + '.join(sorted(subs))})" if subs else name
            lines.append(f"{label} [{gap}]")

        self.add(row, wir_no, Severity.WARN, "Previous activity", "\n".join(lines))

    def check_already_approved(self, row: int, wir_no: str, package: str, building: str,
                               floor: str, area: str, activity: str, index: _Index) -> None:
        tokens = [t for t in tokenize(area)]

        if not tokens:
            entry = _covering(index, package, building, floor, activity, "", self.ref,
                              skip_number=wir_no)
            if entry is not None:
                self.add(row, wir_no, Severity.WARN, "Area submitted before",
                         f"Activity {activity} is already approved at this location "
                         f"as {entry.wir_no}.")
            return

        by_wir: dict[str, list[str]] = defaultdict(list)
        for token in tokens:
            entry = _covering(index, package, building, floor, activity, match_key(token),
                              self.ref, skip_number=wir_no)
            if entry is not None:
                by_wir[entry.wir_no].append(token)

        if by_wir:
            parts = [f"{', '.join(areas)} by {no}" for no, areas in by_wir.items()]
            self.add(row, wir_no, Severity.WARN, "Area submitted before",
                     f"{activity} already approved here for {'; '.join(parts)}.")

    def check_apartments(self, row: int, wir_no: str, package: str, building: str,
                         floor: str, area: str) -> None:
        code = f"{package}-{building}-{floor}".upper()
        known = self.ref.apartments.get(code)
        if not known:
            # Plenty of floors carry no unit list at all. Silence there is
            # correct - only check where there is something to check against.
            return

        unknown = [
            match_key(t) for t in tokenize(area)
            if is_unit_code(match_key(t)) and match_key(t) not in known
        ]
        if unknown:
            self.add(row, wir_no, Severity.WARN, "Apartment",
                     f"{', '.join(unknown)} not in the register for {code}.")

    def check_area_names(self, row: int, wir_no: str, area: str, floor: str) -> None:
        if not self.ref.area_names:
            return

        unknown: list[str] = []
        wrong_floor: list[str] = []

        for token in tokenize(area):
            key = match_key(token)
            if not key or is_unit_code(key) or key.startswith("PART"):
                continue

            canonical = self.ref.canonical_area(token)
            if canonical:
                if not self.ref.floor_allowed(canonical, floor):
                    allowed = ", ".join(sorted(self.ref.area_floors[canonical]))
                    wrong_floor.append(f"{canonical} (only on {allowed})")
            else:
                suggestion = nearest(token, self.ref.area_names)
                unknown.append(
                    f"{token} (did you mean {suggestion}?)" if suggestion else token
                )

        if unknown:
            self.add(row, wir_no, Severity.WARN, "Area name",
                     f"Not a known area name: {', '.join(unknown)}. "
                     f"Add it to {config.NAME_AREA_NAMES} if it is correct.")
        if wrong_floor:
            self.add(row, wir_no, Severity.WARN, "Area name",
                     f"Not on this floor ({floor}): {', '.join(wrong_floor)}.")

    def check_attachments(self, row: int, wir_no: str, values: list) -> None:
        for column, label, warn in ATTACHMENT_COLUMNS:
            raw = text_of(values[column - 1])
            if not raw or raw.startswith("#"):
                continue

            stats = self.attachments[label]
            stats[1] += 1                       # rows where a path was given

            missing: list[str] = []
            resolved = False

            for original in split_paths(raw):
                bare = trim_slashes(unquote(original))
                if exists(bare):
                    resolved = True
                    continue
                fixed = repair(original)
                if fixed != bare and exists(fixed):
                    resolved = True
                else:
                    missing.append(bare)

            if resolved:
                stats[0] += 1                   # rows that resolved

            if warn and missing:
                self.add(row, wir_no, Severity.WARN, "Attachment",
                         f"{label} - not found: {'; '.join(missing)}")


def run(book: LiveWorkbook, ref: ReferenceData) -> RunResult:
    from ..report import write_report          # local import: avoids a cycle

    result = RunResult(TITLE)
    sheet = book.log_sheet

    last_row = book.last_row(sheet, config.Main.WIR_NO)
    if last_row < config.FIRST_DATA_ROW:
        return result.fail(f"No rows found in column B of '{sheet.Name}'.")

    if not ref.previous:
        return result.fail(
            f"Table '{config.TABLE_PREV_ACTIVITIES}' not found in this workbook. "
            "It holds each activity and the activities that must precede it."
        )

    register_sheet = book.sheet(config.REGISTER_SHEET)
    register_rows: list[list] = []
    if register_sheet is not None:
        register_last = book.last_row(register_sheet, config.Register.WIR_NO)
        register_rows = book.block(register_sheet, config.REGISTER_FIRST_ROW,
                                   register_last, config.Register.LAST_COL)

    index = _build_index(register_rows, ref)
    rows = book.block(sheet, config.FIRST_DATA_ROW, last_row, config.Main.LAST_COL)

    state = _Run(book, ref)
    seen: dict[str, int] = {}
    checked = skipped_done = 0

    for offset, values in enumerate(rows):
        row_no = config.FIRST_DATA_ROW + offset
        wir_no = text_of(values[config.Main.WIR_NO - 1])
        if not wir_no:
            continue

        if text_of(values[config.Main.STATUS - 1]).lower() == config.STATUS_DONE.lower():
            skipped_done += 1
            continue

        checked += 1
        dep = text_of(values[config.Main.DEP - 1])
        package = text_of(values[config.Main.PACKAGE - 1])
        building = text_of(values[config.Main.BUILDING - 1])
        floor = text_of(values[config.Main.FLOOR - 1])
        area = normalize_spaces(text_of(values[config.Main.AREA - 1]))
        activity = text_of(values[config.Main.ACTIVITY - 1])
        form = text_of(values[config.Main.FORM - 1])
        description = text_of(values[config.Main.DESCRIPTION - 1])
        rev_text = text_of(values[config.Main.REV - 1])

        state.row_context[row_no] = (dep, package, building, floor, area, activity)

        state.check_number(row_no, wir_no, package, dep, seen)

        rev = 0
        if rev_text.replace(".", "").isdigit():
            rev = int(float(rev_text))
        else:
            state.add(row_no, wir_no, Severity.ERROR, "Rev",
                      f"Column O is '{rev_text}' - must be a number.")

        state.check_already_submitted(row_no, wir_no, rev, index)

        empty = [name for name, value in (
            ("Dep.", dep), ("Package", package), ("Building", building),
            ("Floor", floor), ("Activity", activity)) if not value]
        if empty:
            state.add(row_no, wir_no, Severity.ERROR, "Missing fields",
                      f"Empty: {', '.join(empty)}.")

        if not form:
            state.add(row_no, wir_no, Severity.ERROR, "Checklist",
                      f"Column P is empty - the activity '{activity}' "
                      "did not resolve to a form.")
        elif not book.sheet_exists(form):
            state.add(row_no, wir_no, Severity.ERROR, "Checklist",
                      f"Sheet '{form}' is not in this workbook.")

        if not description:
            state.add(row_no, wir_no, Severity.WARN, "Description",
                      "Column C is empty - the cover sheet will have no description.")

        if floor_is_checkable(floor):
            state.check_apartments(row_no, wir_no, package, building, floor, area)
            state.check_area_names(row_no, wir_no, area, floor)

        if activity and package and building and floor:
            state.check_precedence(row_no, wir_no, package, building, floor,
                                   area, activity, index)
            state.check_already_approved(row_no, wir_no, package, building, floor,
                                         area, activity, index)

        state.check_attachments(row_no, wir_no, values)

    _highlight(book, sheet, last_row, state)
    write_report(book, sheet.Name, checked, skipped_done, state)

    result.findings = state.findings
    result.line(f"Rows checked:        {checked}")
    result.line(f"Rows skipped (Done): {skipped_done}")
    result.line(f"Errors:              {result.errors}")
    result.line(f"Warnings:            {result.warnings}")
    result.line("")
    if result.errors:
        result.line("Fix the errors before generating.")
    elif result.warnings:
        result.line("No errors. Review the warnings, then generate.")
    else:
        result.line("All clear - safe to generate.")
    return result


def _highlight(book: LiveWorkbook, sheet, last_row: int, state: _Run) -> None:
    """Only columns B and I are ever touched, and both are cleared first,
    so this cannot disturb the attachment colouring in U / W / Y / AA."""
    book.clear_fill(sheet, config.Main.WIR_NO, config.FIRST_DATA_ROW, last_row)
    book.clear_fill(sheet, config.Main.AREA, config.FIRST_DATA_ROW, last_row)

    for row, severity in state.row_severity.items():
        colour = config.COLOUR_ERROR if severity is Severity.ERROR else config.COLOUR_WARN
        book.fill(sheet, row, config.Main.WIR_NO, colour)

    for row, severity in state.area_severity.items():
        colour = config.COLOUR_ERROR if severity is Severity.ERROR else config.COLOUR_WARN
        book.fill(sheet, row, config.Main.AREA, colour)
