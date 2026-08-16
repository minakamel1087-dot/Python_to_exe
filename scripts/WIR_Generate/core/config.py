"""
Every layout fact about the workbook lives here.

Nothing else in the project hardcodes a column letter or a sheet name, so
when the sheet moves again - and it has, twice - this is the only file
that changes.
"""

from __future__ import annotations

# --- The workbook ----------------------------------------------------------

WORKBOOK_NAME = "i125-WIR Cover Generate.xlsm"

# Sheets that are never the WIR log, so the tools refuse to run on them.
NON_LOG_SHEETS = {
    "WIR-Form", "PreFlight", "WIRs", "WIR_History", "WIR_Comments",
    "Comments_Form", "Project-Des.", "ELE_Data", "PUB_Data", "MacroLog",
}

REPORT_SHEET = "PreFlight"
REGISTER_SHEET = "WIRs"


# --- The log sheet (Main) --------------------------------------------------
# Headers sit on row 2, data starts on row 3.

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
    ATT_PREV = 25     # Y
    ATT_FIRST = 27    # AA
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
    PATH = 14         # N  commented-WIR path
    REV = 15          # O
    STATUS = 16       # P
    LAST_COL = 30     # AD


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


# --- Colours (BGR longs, the order Excel wants) ----------------------------

def _bgr(r: int, g: int, b: int) -> int:
    return r + (g << 8) + (b << 16)

COLOUR_ERROR = _bgr(255, 199, 206)   # soft red
COLOUR_WARN = _bgr(255, 235, 156)    # amber
COLOUR_NONE = -4142                  # xlNone
