"""
Every layout fact about the workbook lives here.

Nothing else in the project hardcodes a column letter or a sheet name, so
when the sheet moves again - and it has, twice - this is the only file
that changes.
"""

from __future__ import annotations

import os

# --- The workbook ----------------------------------------------------------

WORKBOOK_NAME = "i125-WIR Cover Generate.xlsm"

# Sheets that are never the WIR log, so the tools refuse to run on them.
NON_LOG_SHEETS = {
    "WIR-Form", "PreFlight", "WIRs", "WIR_History", "WIR_Comments",
    "Comments_Form", "Project-Des.", "ELE_Data", "PUB_Data", "MacroLog",
    "Paths",
}

REPORT_SHEET = "PreFlight"
REGISTER_SHEET = "WIRs"

# --- Server / local paths --------------------------------------------------
# The pairs live in the workbook so each person can point the local column
# at their own PC. See core/pathmap.py.

PATHS_SHEET = "Paths"


class Paths:
    ITEM = 1          # A
    SERVER = 2        # B
    LOCAL = 3         # C
    FIRST_ROW = 2     # row 1 is the header


# Which set is in use, as a number the sheet's own formulas read:
#   =SWITCH($Y$1, 1, ELE_Server_path, 2, ELE_Local_path)
# The program writes it; it is reset to Server on every start.
MODE_SHEET = "Main"   # where the SWITCH() formulas read it
MODE_ROW = 1
MODE_COL = 25         # Y


# --- The log sheet (Main) --------------------------------------------------
# Headers sit on row 2, data starts on row 3. Row 1 is left blank, and
# WIR_History is laid out the same way - anything that reads or writes a
# header must use HEADER_ROW rather than assuming row 1.

HEADER_ROW = 2
FIRST_DATA_ROW = 3

class Main:
    WIR_NO = 2        # B  Submittal Reference. No.
    DESCRIPTION = 3   # C
    EXCLUSION = 4     # D
    DEP = 5           # E  trade
    PACKAGE = 6       # F
    BUILDING = 7      # G
    FLOOR = 8         # H
    AREA = 9          # I
    ACTIVITY = 10     # J
    SUGGEST = 11      # K  Check Areas writes here, never to AREA
    APARTMENTS = 12   # L
    LOCATION = 13     # M
    LINK = 14         # N  the typed attachment path
    REV = 15          # O
    FORM = 16         # P  checklist sheet name
    DATE_SUBMIT = 19  # S
    DATE_INSPECT = 20 # T
    ATT_SITE = 21     # U
    ATT_QC = 23       # W

    # The two names below read backwards, and the sheet headers are the
    # authority. Check the column letter, not the constant name.
    #
    #   Y  (ATT_PREV)  "Previous WIR Rev." - the earlier revision of this
    #                  same WIR, attached when Rev in column O is above 0.
    #                  Copied whole.
    #   AA (ATT_FIRST) "Previous Activity" - the approved predecessor
    #                  WIRs, filled by Import Previous WIR Paths. These
    #                  are full commented WIRs and Generate takes only
    #                  their first page.
    ATT_PREV = 25     # Y   Previous WIR Rev.
    ATT_FIRST = 27    # AA  Previous Activity
    STATUS = 29       # AC
    LAST_COL = 29


# --- The register (WIRs sheet) --------------------------------------------
# Headers on row 3, data from row 4.

REGISTER_FIRST_ROW = 4

class Register:
    WIR_NO = 2        # B
    PACKAGE = 6       # F
    BUILDING = 7      # G
    FLOOR = 8         # H
    AREA = 9          # I
    ACTIVITY = 10     # J
    PATH = 14         # N   commented-WIR path, as the register stores it
    REV = 15          # O
    STATUS = 16       # P
    # AD holds the same file under the local extract folder. That is the
    # one that actually opens, so it is what gets copied into the log.
    LOCAL_PATH = 30   # AD
    LAST_COL = 30


# --- Reference tables, found by name rather than position ------------------

TABLE_PREV_ACTIVITIES = "Prev_Activities"     # Excel Table
TABLE_FLOOR_APARTMENTS = "Floor_Apartments"   # Excel Table
NAME_ACTIVITY_SPLITS = "Activity_Splits"      # defined name
NAME_STATUS_CODES = "WIR_Status_Codes"        # defined name
NAME_AREA_NAMES = "Area_Names"                # defined name

# Floor_Apartments column positions inside the table
REG_CODE = 0
REG_APT_TEXT = 1
REG_APT_FIRST = 5    # Apt01
REG_APT_LAST = 19    # Apt15


# --- Behaviour -------------------------------------------------------------

STATUS_DONE = "Done"
WIR_NO_LENGTH = 42
DEFAULT_APPROVED_CODE = "B"

# Floors with a defined apartment band and room scheme. Anything else
# (SC01, SC02, RS, PD, UR, or a typo like "1F-2F") is not area-checked -
# there is no agreed list to check it against.
CHECKED_FLOORS = {
    "BS", "GF", "1F", "2F", "3F", "4F", "5F",
    "6F", "7F", "8F", "9F", "10F", "RF",
}

# Systems rather than places. They turn up in the Area column on a scope-
# of-work row and are accepted without comment - they are not areas, and
# reporting them as unknown ones is noise.
#
# BMS is deliberately absent: BMS Room is a real room elsewhere in the log,
# and Area_Names already carries it.
SYSTEM_WORDS = {
    "ACCESS CONTROL", "ACCES CONTROL", "CCTV", "FA VOICE EVACUATION",
    "EM LIGHTING", "LIGHTING", "POWER",
}

# An Area that describes the extent of the work rather than naming a place
# - "FULL SLAB", "GF to First Floor columns", "Points Excluded From
# WIR-000927". Free text is legitimate here, so a value carrying one of
# these words is accepted silently rather than reported as unknown.
FREE_TEXT_MARKERS = {
    "HIGHLIGHTED", "EXCLUDED", "SLAB", "RAFT", "COLUMN", "COLUMNS",
    "PARAPET", "CONDUCTOR", "WALL", "WALLS", "DRAWING",
}

# Accepted in the Area column without comment and without a suggestion.
# Matched against match_key, which has already dropped a trailing ROOM or
# ROOMS - so "Service rooms" arrives here as "SERVICE", while "Service
# corridor" does not and stays a real area of its own.
IGNORED_AREAS = {
    "LIFT CORE", "LIFTCORE", "SERVICE", "ALL SERVICES", "SERVICES",
    # Mock-up is free text: not an area, and not a prefix on one either.
    "MOCK UP", "MOCKUP",
}

# A bare "Toilet" does not say whose. Male Toilets and Female Toilets are
# separate areas, so the answer cannot be guessed - the note asks.
AMBIGUOUS_AREAS = {
    "TOILET": "Male Toilets, Female Toilets, or both?",
    "TOILETS": "Male Toilets, Female Toilets, or both?",
    "BATHROOM": "Male Toilets, Female Toilets, or both?",
    "WASHROOM": "Male Toilets, Female Toilets, or both?",
}

# A leftover fragment this short says nothing useful - it is what is left
# after a split, not a name someone typed. Reporting it only adds noise.
MIN_REPORTABLE_AREA = 4

# --- Expiry check ----------------------------------------------------------
# A published text file holding the date this program stops working. Empty
# switches the check off entirely. See core/licence.py for the format and
# for what happens with no network.
# Checked in this order: the local file first (network never touched when
# it is there), then the published one. Both must carry a valid signature.
# Per-user, and built from the environment rather than spelled out: a
# literal path with a user name in it would only be right on one machine.
LICENCE_LOCAL_FILE = os.path.join(
    os.environ.get("LOCALAPPDATA") or os.environ.get("TEMP") or ".",
    "WGT", "i125_WIR_Validation.txt",
)
LICENCE_URL = (
    "https://raw.githubusercontent.com/minakamel1087-dot/"
    "Document-management-system-date-validation/main/i125_WIR_Generate"
)
LICENCE_TIMEOUT = 4.0        # seconds; startup must not hang on a dead host

# The public half of the signing key. Verifies signatures, cannot create
# them, and is safe to publish. Empty switches the whole check off.
# Produced by: python tools\sign_licence.py --new-key
LICENCE_PUBLIC_KEY = "8NkMdB06R1DZ5iUPodMEyqysuL7Jblv66Qjv+Y6G9w8="


# Where site attachments actually live.
SHARE_ROOT = "Z:\\Common\\"

# Drive letters that mean a real folder on this PC rather than the share
# under a different letter. These are never rewritten.
LOCAL_DRIVES = set("CDEFGH")

# Fix WIR Path Prefix: the register stores the commented WIRs on the
# network share; the local extract folder is what actually gets opened.
REGISTER_NETWORK_PREFIX = (
    "\\\\192.168.225.6\\i125MEP\\QC\\1. Work Inspection Request\\Commented WIR\\ELE\\"
)
REGISTER_LOCAL_PREFIX = "D:\\i125\\Logs\\WIR\\ELE&ELV\\extracted\\"

# Extract WIR Cover Page. A standalone tool: it reads no workbook and
# writes none, so it runs whether or not Excel is open.
#
# The destination is the same folder Fix WIR Path Prefix points the
# register at - these covers are what those links open.
COVER_SOURCE = "Z:\\QC\\1. Work Inspection Request\\Commented WIR\\ELE&ELV"
COVER_DEST = REGISTER_LOCAL_PREFIX


# --- Colours (BGR longs, the order Excel wants) ----------------------------

def _bgr(r: int, g: int, b: int) -> int:
    return r + (g << 8) + (b << 16)

COLOUR_ERROR = _bgr(255, 199, 206)   # soft red
COLOUR_WARN = _bgr(255, 235, 156)    # amber
COLOUR_NONE = -4142                  # xlNone
