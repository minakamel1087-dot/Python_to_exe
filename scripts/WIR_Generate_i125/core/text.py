"""
Turning what people type in the Area column into something comparable.

The rules here are not guesses - each one exists because the register
contains data that broke the previous version:

  non-breaking spaces   "Apartment<nbsp>201" matched nothing at all
  glued unit numbers    Excel stores "101,102,103" as the number 101102103
  missing separators    "301 302 303 304", "Garbage Room and Water Meter Room"
  misspelt prefixes     "Apartmants 114", "Flats 301"
  stray brackets        "(Staircase 1, 2 and Lift Core)"
"""

from __future__ import annotations

import re

NBSP = " "

# Words that introduce a unit number rather than naming a place. The APART
# and FLAT tests are prefix matches so the misspellings in the log
# ("Apartmants", "Aparments") are caught too.
_UNIT_WORDS = {"APT", "UNIT", "NO", "PACKAGE", "BUILDING"}

# A dot separates in two shapes: "EN. Facility room" (dot then space) and
# "G01.G02" (no spaces at all, a comma someone missed).
#
# It is deliberately not treated as noise everywhere. "Garbage .R" needs
# its dot to expand to Room, and "1.5" is one number - hence the guard
# against a digit on both sides.
_SEPARATORS = re.compile(
    r"[,;/&]| and |\.\s+|(?<!\d)\.(?=\w)|(?<=\w)\.(?!\d)",
    re.IGNORECASE,
)

# What is left clinging to a token once the separators have done their
# work: "Substation South and" when the Building reference after it was
# removed, or a stray bracket, comma or dash.
_EDGE_NOISE = re.compile(r"^(?:and\s+|&\s*)|(?:\s+and|\s*&)$", re.IGNORECASE)
_LEADING_LETTERS = re.compile(r"^[A-Z]+")

# A hyphen separates two areas - "Telephone-Water Meter Room", "Corridor -
# Store" - except when a digit sits on either side of it. That exception is
# what protects "Part-5" and "A1-004", which are one thing and not two.
_AREA_HYPHEN = re.compile(r"(?<!\d)\s*-\s*(?!\d)")

# Package and Building belong in their own columns. Where they have been
# typed into the Area as well they are noise, and they push the parts of
# the value that do matter out of reach.
_SCOPE_REFERENCE = re.compile(
    r"\b(?:PACKAGE|PKG)\s*\.?\s*\d+\b"
    r"|\bBUILDING\s+[A-Z]?\d+(?:\s*-\s*\d+)?\b",
    re.IGNORECASE,
)


def strip_scope_references(text: str) -> str:
    """Removes "Package 09" and "Building A1 - 004" from an Area value."""
    return normalize_spaces(_SCOPE_REFERENCE.sub(" ", text)).strip()


# "Staircase 1", "staircase 01", "Stair case 2", "Staircase 1 (Basement to
# 5th Floor)". The number is what says which staircase, and it belongs in
# the Floor column as SC01 or SC02.
_STAIRCASE = re.compile(r"\bSTAIR\s?CASE\b[\s:.-]*0*(\d+)?", re.IGNORECASE)


def staircase_number(token: str) -> tuple[bool, str]:
    """(is this a staircase, which one). The number comes back without its
    leading zeros, so "staircase 01" and "Staircase 1" are one thing."""
    match = _STAIRCASE.search(token)
    if not match:
        return False, ""
    return True, match.group(1) or ""


# Mock-up is a prefix on a real area, not an area itself: "Mock up
# Electrical Room" is the electrical room, built as a mock-up. Stripping it
# lets the rest be recognised, and it is put back in the suggestion.
_MOCKUP = re.compile(r"^\s*MOCK\s*[-\s]?\s*UP\b[\s:.,-]*", re.IGNORECASE)
MOCKUP_PREFIX = "Mock up"

# Settled before the hyphen splitter runs, or "Mock-up" comes apart into
# "Mock" and "up" and neither half means anything.
_MOCKUP_HYPHEN = re.compile(r"\bMOCK\s*-\s*UP\b", re.IGNORECASE)


def split_mockup(token: str) -> tuple[bool, str]:
    """(was it prefixed, what is left). "Mock up" on its own leaves ""."""
    match = _MOCKUP.match(token)
    if not match:
        return False, token
    return True, token[match.end():].strip()


# Ground-floor units are G01..G11 in the register, but the log also writes
# them GF01, GF1 and "GF 02". Normalised to the register's form, and only
# where a number follows - a bare "GF" is the floor code and must survive
# untouched, as in "GF to First Floor columns".
_GF_UNIT = re.compile(r"\bGF\s*0*(\d{1,2})\b", re.IGNORECASE)


def _to_g_unit(match: re.Match) -> str:
    return f"G{int(match.group(1)):02d}"


# "Garbage .R" is Garbage Room. Expanded rather than added as a variant,
# because the abbreviation is a habit and not specific to one room - and
# widening the fuzzy-match threshold to reach it is what silently truncated
# "Substation South _Attic" once before.
_ROOM_ABBREV = re.compile(r"\s*\.\s*R\b\.?", re.IGNORECASE)


# "2nd to 3rd", "first to 2nd", "Basement to Roof Floor" - the extent of
# the work rather than a place, same as the other free text.
_FLOOR_RANGE = re.compile(
    r"\b(?:BS|GF|RF|BASEMENT|ROOF|GROUND|FIRST|SECOND|THIRD|FOURTH|FIFTH"
    r"|\d+\s*(?:ST|ND|RD|TH)?)\s+TO\s+",
    re.IGNORECASE,
)


def is_floor_range(token: str) -> bool:
    return bool(_FLOOR_RANGE.search(token))


def normalize_spaces(text: str) -> str:
    """Non-breaking spaces, tabs and line breaks all become plain spaces,
    and runs of spaces collapse. Cells pasted from Word, Outlook or a web
    page carry these, and str.strip() does not remove them."""
    if not text:
        return ""
    out = text.replace(NBSP, " ").replace("\t", " ")
    out = out.replace("\r", " ").replace("\n", " ")
    while "  " in out:
        out = out.replace("  ", " ")
    return out


def tokenize(area: str) -> list[str]:
    """Splits an Area value into its parts, exactly as typed.

    Nothing is normalised here - the report quotes these back, so
    "Garbage Room" has to stay "Garbage Room" and not become "Garbage".
    """
    if not area:
        return []

    text = normalize_spaces(area).replace("(", " ").replace(")", " ")
    text = strip_scope_references(text)
    text = _MOCKUP_HYPHEN.sub(MOCKUP_PREFIX, text)
    text = _ROOM_ABBREV.sub(" Room", text)
    text = _GF_UNIT.sub(_to_g_unit, text)
    text = _AREA_HYPHEN.sub(", ", text)
    parts = [_EDGE_NOISE.sub("", p.strip()).strip(" .,-") for p in _SEPARATORS.split(text)]
    tokens = [p for p in parts if p and p.upper() != "MISSING"]
    return _inherit_names(_expand_number_runs(tokens))


_NAMED_NUMBER = re.compile(r"^(?P<stem>\D+?)\s+\d+$")


def _inherit_names(tokens: list[str]) -> list[str]:
    """"Staircase 1, 2" means staircase one and staircase two.

    A bare number following a "<name> <number>" token takes that name.
    Only applies when the previous token has a non-numeric stem and a
    space before its number, so "301, 302" and "G01, G02" are untouched -
    those are already complete on their own.
    """
    out: list[str] = []
    stem = ""

    for token in tokens:
        match = _NAMED_NUMBER.match(token)
        if match:
            stem = match.group("stem").strip()
            out.append(token)
        elif stem and token.isdigit():
            out.append(f"{stem} {token}")
        else:
            if not token.isdigit():
                stem = ""
            out.append(token)

    return out


def _expand_number_runs(tokens: list[str]) -> list[str]:
    """"301 302 303 304" and "Apartments 201 202" arrive as one token
    because nobody typed the commas. A token is split only when *every*
    part is a unit code - "G10 Kitchens" has a room name in it and has to
    stay whole."""
    out: list[str] = []
    for token in tokens:
        stripped = match_key(token)
        parts = stripped.split(" ")
        if len(parts) > 1 and all(is_unit_code(p) for p in parts):
            out.extend(parts)
        else:
            out.append(token)
    return out


def is_unit_word(stem: str) -> bool:
    """Does this leading word mean "apartment"?

    Matched loosely, because the log contains "Apartmants", "Aparments"
    and "Aartments" - all plainly meant to introduce a unit number. A
    strict test leaves those looking like room names, and the number
    inside them never gets checked against the register.
    """
    word = stem.upper()
    if not word:
        return False
    if word.startswith("APART") or word.startswith("FLAT") or word in _UNIT_WORDS:
        return True
    if len(word) >= 4:
        for target in ("APARTMENT", "APARTMENTS"):
            if edit_distance(word, target) <= 2:
                return True
    return False


def match_key(token: str) -> str:
    """The comparable form of one token: unit-word prefix removed,
    trailing ROOM/ROOMS removed, upper-cased.

    Only the leading run of *letters* is treated as the prefix, never the
    first whole word - the log contains both "Apartment 201" and
    "Apartment201", and dropping the whole word would swallow the number
    in the second so the two would never match each other.
    """
    text = normalize_spaces(token).strip().upper()
    if not text:
        return ""

    if text == "PART" or text.startswith("PART "):
        return text

    # Repeatedly, not once: the log writes "Apartment no 401" and
    # "Apartments no. 101", which is two unit words in a row. Stripping
    # only the first leaves "NO 401" - not a unit code, so the number is
    # never checked against the register and the whole value reads as an
    # unknown room name.
    while True:
        match = _LEADING_LETTERS.match(text)
        stem = match.group(0) if match else ""
        if not stem or not is_unit_word(stem):
            break
        text = text[len(stem):].strip().lstrip(":.-").strip()

    for suffix in (" ROOMS", " ROOM"):
        if text.endswith(suffix):
            text = text[: -len(suffix)].strip()
            break

    return text


def is_unit_code(token: str) -> bool:
    """Looks like an apartment: has a digit and no space. Distinguishes
    "204" and "G01" from "Telephone Room"."""
    return bool(token) and any(c.isdigit() for c in token) and " " not in token


def has_unit_prefix(raw_token: str) -> bool:
    """Was the number introduced by a word that means "apartment"?

    "Flats 301" says so outright, and so does "Aartments 206" once the
    typo is allowed for. A bare "2" does not, and "Staircase 1, 2" is
    exactly where guessing goes wrong.
    """
    text = normalize_spaces(raw_token).strip()
    match = _LEADING_LETTERS.match(text.upper())
    return bool(match) and is_unit_word(match.group(0))


def is_unit_for_floor(key: str, floor: str) -> bool:
    """Does this code fall in the floor's own numbering band?

    GF runs G01..G99; 1F..10F use the floor number followed by two
    digits. Anything else - a bare "2", a "503" on the fourth floor - is
    not this floor's apartment and must not be relabelled as one.
    """
    band = normalize_spaces(floor).strip().upper()
    if band == "GF":
        return len(key) == 3 and key[0] == "G" and key[1:].isdigit()
    if band.endswith("F") and band[:-1].isdigit():
        digits = band[:-1]
        return key.isdigit() and len(key) == len(digits) + 2 and key.startswith(digits)
    return False


def area_key(name: str) -> str:
    """Comparison form of an area *name*: letters and digits only, upper
    case, trailing S dropped so "Store Rooms" and "Store Room" are one
    thing."""
    text = normalize_spaces(name).strip().upper()
    kept = [c for c in text if c.isalnum() or c == " "]
    key = "".join(kept).strip()
    while "  " in key:
        key = key.replace("  ", " ")
    if len(key) > 3 and key.endswith("S"):
        key = key[:-1]
    return key


def split_glued_numbers(token: str, floor: str) -> list[str]:
    """Excel reads "101,102,103" as a number with thousand separators and
    stores 101102103 - the commas are gone before any code sees the cell.

    Recoverable: split into 3-digit codes and accept only if every one of
    them lands in the floor's own band (1F -> 1xx). If any chunk does not
    match, nothing is returned rather than a guess.
    """
    if len(token) <= 3 or len(token) % 3 or not token.isdigit():
        return []

    band = floor.strip().upper()
    if len(band) < 2 or not band.endswith("F"):
        return []
    band = band[:-1]
    if not band.isdigit() or band == "0":
        return []

    chunks = [token[i:i + 3] for i in range(0, len(token), 3)]
    if all(c.startswith(band) for c in chunks):
        return chunks
    return []


def edit_distance(a: str, b: str) -> int:
    """Levenshtein, two rows only - the full matrix is pointless for
    strings this short."""
    if a == b:
        return 0
    if not a:
        return len(b)
    if not b:
        return len(a)

    previous = list(range(len(b) + 1))
    for i, ca in enumerate(a, start=1):
        current = [i]
        for j, cb in enumerate(b, start=1):
            cost = 0 if ca == cb else 1
            current.append(min(current[j - 1] + 1, previous[j] + 1, previous[j - 1] + cost))
        previous = current
    return previous[-1]


# A correction is written straight into the suggestion, so it has to be
# a typo-level difference and nothing more. Two edits catches "Cooridor",
# "Elctrical" and "Garabge"; the length guard stops a longer value being
# truncated to a shorter name it happens to contain - "Substation South
# _Attic" is not "Substation South", it is that plus a room.
_MAX_EDITS = 2
_MAX_LENGTH_GAP = 2

# Two edits is most of a three-letter word, so short tokens match almost
# anything: "GF" came back as "Gym". Abbreviations that do mean a room -
# MTR, BMS, CBS, GSM, RMU - are listed in Area_Names and found by exact
# lookup, so they never needed the fuzzy path.
_MIN_FUZZY_LENGTH = 4


def nearest(name: str, candidates: list[str]) -> str:
    """The closest known name, or "" when nothing is close enough."""
    probe = area_key(name)
    if len(probe) < _MIN_FUZZY_LENGTH:
        return ""

    best, best_distance = "", 99
    for candidate in candidates:
        key = area_key(candidate)
        if abs(len(key) - len(probe)) > _MAX_LENGTH_GAP:
            continue
        distance = edit_distance(probe, key)
        if distance < best_distance:
            best, best_distance = candidate, distance

    return best if best_distance <= _MAX_EDITS else ""
