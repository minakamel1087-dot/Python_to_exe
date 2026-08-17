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

_SEPARATORS = re.compile(r"[,;/&]| and ", re.IGNORECASE)
_LEADING_LETTERS = re.compile(r"^[A-Z]+")


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
    parts = [p.strip() for p in _SEPARATORS.split(text)]
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

    match = _LEADING_LETTERS.match(text)
    stem = match.group(0) if match else ""

    if is_unit_word(stem):
        text = text[len(stem):].strip()
        text = text.lstrip(":.-").strip()

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


def nearest(name: str, candidates: list[str]) -> str:
    """The closest known name, or "" when nothing is close enough."""
    probe = area_key(name)
    if not probe:
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
