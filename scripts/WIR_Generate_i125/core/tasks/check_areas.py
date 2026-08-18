"""
Check Areas - writes into column K and nowhere else.

The Area column itself is never overwritten, so nothing is lost if a
suggestion is wrong: K gets the suggested Area text when the value can be
tidied, and the reason when it cannot. There is no report sheet - the
answer belongs next to the row it is about.

Floors without a defined area scheme are skipped entirely.
"""

from __future__ import annotations

from .. import config
from ..findings import RunResult
from ..reference import ReferenceData
from ..text import (
    has_unit_prefix,
    is_floor_range,
    is_unit_code,
    is_unit_for_floor,
    match_key,
    nearest,
    normalize_spaces,
    split_glued_numbers,
    staircase_number,
    tokenize,
)
from ..workbook import LiveWorkbook, text_of

TITLE = "Check Areas"


def floor_is_checkable(floor: str) -> bool:
    return floor.strip().upper() in config.CHECKED_FLOORS


def _known(token: str, ref: ReferenceData) -> str:
    """The canonical name for this token, by spelling or by near miss."""
    return ref.canonical_area(token) or nearest(token, ref.area_names) or ""


def _ignorable(token: str) -> bool:
    """Accepted without comment: a part reference, a system, free text, a
    floor range, or something deliberately not checked."""
    key = match_key(token)
    return bool(
        key.startswith("PART")
        or key in config.IGNORED_AREAS
        or _is_system(token)
        or _is_free_text(token)
        or is_floor_range(token)
    )


def _unit_with_words(token: str) -> bool:
    """"G10 Kitchens", "Dry Wall G01" - a unit number and a room name in
    one token, with no separator between them.

    Splitting it would have to guess which part is the area, so the value
    is kept exactly as typed and reported for someone to deal with.
    """
    words = token.split()
    if len(words) < 2:
        return False
    units = [w for w in words if is_unit_code(match_key(w))]
    return bool(units) and len(units) < len(words)


def _resolve_one(part: str, ref: ReferenceData) -> str:
    """A known name, or the text itself when it is something we accept as
    written. "" when it is neither."""
    name = _known(part, ref)
    if name:
        return name
    return part if _ignorable(part) else ""


def split_known(token: str, ref: ReferenceData, depth: int = 0) -> list[str]:
    """"Telephone room Garbage room" is two rooms typed without a comma
    between them. Split only where *every* part is recognised - a partial
    match would invent an area that nobody wrote.

    The first cut that works wins, and the tail is split again, so
    "water meter room corridor LL" comes apart into both of its rooms.
    """
    words = token.split()
    if len(words) < 2 or depth > 3:
        return []

    for cut in range(1, len(words)):
        left = _resolve_one(" ".join(words[:cut]), ref)
        if not left:
            continue
        rest = " ".join(words[cut:])
        right = _resolve_one(rest, ref)
        if right:
            return [left, right]
        deeper = split_known(rest, ref, depth + 1)
        if deeper:
            return [left] + deeper

    return []


def format_area(area: str, floor: str, ref: ReferenceData) -> str:
    """Rewrites one Area value into the house format:

        "301 302 303 304"        -> "Apartments 301, 302, 303, 304"
        "101,102,103"            -> "Apartments 101, 102, 103"
        "Cooridor"               -> "Corridor"
        "MTR"                    -> whatever Area_Names calls it

    Unit codes are collected under one "Apartments" prefix in the order
    they were typed - not sorted, because the typed order sometimes
    carries meaning and reordering it silently would be surprising.
    Anything unknown is left exactly as written.
    """
    if not area.strip():
        return area

    units: list[str] = []
    names: list[str] = []
    seen: set[str] = set()

    for token in tokenize(area):
        key = match_key(token)
        if not key:
            continue

        # Recover a comma list Excel turned into one long number. This
        # already proves the band itself - every chunk has to land in it.
        glued = split_glued_numbers(key, floor)

        # Otherwise only call it an apartment when the text says so or the
        # number belongs to this floor's band. Without that, "Staircase 1,
        # 2" becomes "Apartments 2", and accepting that suggestion would
        # put a wrong area on the WIR.
        claimed = bool(glued) or has_unit_prefix(token) or is_unit_for_floor(key, floor)

        if is_unit_code(key) and claimed:
            for unit in glued or [key]:
                if unit not in seen:
                    seen.add(unit)
                    units.append(unit)
        elif key.startswith("PART"):
            if token not in seen:
                seen.add(token)
                names.append(token)
        elif _ignorable(token) or _unit_with_words(token):
            # Free text, a system, a floor range. Suggesting anything here
            # is worse than suggesting nothing: "GF to First Floor columns"
            # came back as "Gym, to First Floor columns".
            if token not in seen:
                seen.add(token)
                names.append(token)
        else:
            # Known spelling wins, then the closest known name, then two
            # names typed without a separator between them. The suggestion
            # carries the correction rather than asking a question, and
            # anything still unrecognised stays exactly as typed.
            canonical = _known(token, ref)
            found = [canonical] if canonical else split_known(token, ref) or [token]
            for name in found:
                if name not in seen:
                    seen.add(name)
                    names.append(name)

    parts: list[str] = []
    if units:
        parts.append("Apartments " + ", ".join(units))
    if names:
        parts.append(", ".join(names))

    return ", ".join(parts) if parts else area


def _is_free_text(token: str) -> bool:
    """Does this describe the extent of the work rather than name a place?"""
    words = set(token.upper().replace("-", " ").split())
    return bool(words & config.FREE_TEXT_MARKERS)


def _is_system(token: str) -> bool:
    return token.strip().upper() in config.SYSTEM_WORDS


def area_comment(area: str, floor: str, package: str, building: str,
                 ref: ReferenceData) -> str:
    """Everything about this row's Area that could not be resolved."""
    if not area.strip():
        # A blank Area means the whole floor. That is legitimate, so it
        # gets no comment.
        return ""

    code = f"{package}-{building}-{floor}".upper()
    known = ref.apartments.get(code, set())
    notes: list[str] = []

    for token in tokenize(area):
        key = match_key(token)
        if not key:
            continue

        question = config.AMBIGUOUS_AREAS.get(key)
        if question:
            notes.append(f"{token}: {question}")
            continue

        # A staircase is identified by its number, and that belongs in the
        # Floor column - so here it is the Floor that needs correcting,
        # not the Area.
        is_stair, number = staircase_number(token)
        if is_stair:
            if not number:
                notes.append(
                    f"{token}: which staircase? The number belongs in the "
                    "Floor column as SC01 or SC02."
                )
            elif floor.strip().upper() != f"SC{int(number):02d}":
                notes.append(
                    f"{token}: Floor should be SC{int(number):02d}, not {floor}."
                )
            continue

        if is_unit_code(key):
            if known and key not in known and not split_glued_numbers(key, floor):
                notes.append(f"{key} is not in the register for {code}.")
            continue

        if _ignorable(token):
            continue

        if _unit_with_words(token):
            notes.append(
                f"{token}: a unit number and a room name together - "
                "left as typed."
            )
            continue

        canonical = _known(token, ref)
        if canonical:
            if not ref.floor_allowed(canonical, floor):
                allowed = ", ".join(sorted(ref.area_floors[canonical]))
                notes.append(f"{canonical} is only on {allowed}, not {floor}.")
        elif not split_known(token, ref):
            # A close match, or a pair of names typed without a separator,
            # is already applied in the suggestion - saying "did you mean"
            # as well is just noise. Only what cannot be resolved needs a
            # note, and not a fragment left over from a split.
            if len(key) >= config.MIN_REPORTABLE_AREA:
                notes.append(f"{token} is not a known area.")

    return " ".join(notes)


def run(book: LiveWorkbook, ref: ReferenceData) -> RunResult:
    result = RunResult(TITLE)
    sheet = book.log_sheet

    last_row = book.last_row(sheet, config.Main.WIR_NO)
    if last_row < config.FIRST_DATA_ROW:
        return result.fail(f"No rows found in column B of '{sheet.Name}'.")

    rows = book.block(sheet, config.FIRST_DATA_ROW, last_row, config.Main.LAST_COL)

    # Last run's output goes first, or a fixed row keeps its old note.
    count = last_row - config.FIRST_DATA_ROW + 1
    book.write_column(sheet, config.Main.SUGGEST, config.FIRST_DATA_ROW, [""] * count)
    book.clear_fill(sheet, config.Main.SUGGEST, config.FIRST_DATA_ROW, last_row)

    column: list[str] = []
    checked = suggested = commented = skipped_floor = skipped_done = 0
    flag_rows: list[int] = []

    for offset, values in enumerate(rows):
        row_no = config.FIRST_DATA_ROW + offset
        column.append("")

        wir_no = text_of(values[config.Main.WIR_NO - 1])
        if not wir_no:
            continue

        if text_of(values[config.Main.STATUS - 1]).lower() == config.STATUS_DONE.lower():
            skipped_done += 1
            continue

        floor = text_of(values[config.Main.FLOOR - 1])
        if not floor_is_checkable(floor):
            skipped_floor += 1
            continue

        checked += 1
        area = normalize_spaces(text_of(values[config.Main.AREA - 1]))
        package = text_of(values[config.Main.PACKAGE - 1])
        building = text_of(values[config.Main.BUILDING - 1])

        suggestion = format_area(area, floor, ref) if area else ""
        if suggestion == area:
            suggestion = ""

        comment = area_comment(area, floor, package, building, ref)

        cell = suggestion
        if comment:
            cell = f"{cell}\n{comment}" if cell else comment

        column[-1] = cell
        if comment:
            commented += 1
            flag_rows.append(row_no)
        elif cell:
            suggested += 1

    book.write_column(sheet, config.Main.SUGGEST, config.FIRST_DATA_ROW, column)
    for row_no in flag_rows:
        book.fill(sheet, row_no, config.Main.SUGGEST, config.COLOUR_WARN)

    sheet.Columns(config.Main.SUGGEST).WrapText = True
    sheet.Columns(config.Main.SUGGEST).VerticalAlignment = -4160   # xlTop

    result.line(f"Rows checked:          {checked}")
    result.line(f"Suggested corrections: {suggested}")
    result.line(f"Needs a look:          {commented}")
    result.line(f"Skipped (floor):       {skipped_floor}")
    result.line(f"Skipped (Done):        {skipped_done}")
    result.line("")
    result.line("Results are in column K. The Area column is untouched.")
    return result
