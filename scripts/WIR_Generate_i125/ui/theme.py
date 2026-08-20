"""
Colours and the stylesheet.

Everything the MSForms version could not do - rounded corners, real hover
states, proper spacing - is a line of Qt stylesheet here.
"""

from __future__ import annotations

# Section accents
TEAL = "#1D9E75"
PURPLE = "#534AB7"
CORAL = "#D85A30"

TEAL_TINT = "#E1F5EE"
PURPLE_TINT = "#EEEDFE"
CORAL_TINT = "#FAECE7"

TEAL_TEXT = "#0F6E56"
PURPLE_TEXT = "#3C3489"
CORAL_TEXT = "#993C1D"

# Surfaces
PAGE = "#F5F6F8"
CARD = "#FFFFFF"
BORDER = "#E3E5E9"
TEXT = "#2C2F36"
MUTED = "#7A808C"
NAVY = "#1F4E78"

OK = "#1D9E75"
FAIL = "#C0392B"

SECTIONS = {
    "tools": (TEAL, TEAL_TINT, TEAL_TEXT),
    "verify": (PURPLE, PURPLE_TINT, PURPLE_TEXT),
    "generate": (CORAL, CORAL_TINT, CORAL_TEXT),
}


def stylesheet() -> str:
    return f"""
    QWidget {{
        background: {PAGE};
        color: {TEXT};
        font-family: "Segoe UI";
        font-size: 10pt;
    }}
    QLabel#title {{
        font-size: 17pt;
        font-weight: 600;
        color: {NAVY};
    }}
    QLabel#status {{
        color: {MUTED};
        font-size: 9pt;
    }}
    QLabel#sectionHeader {{
        color: #FFFFFF;
        font-size: 8.5pt;
        font-weight: 600;
        padding: 5px 10px;
        border-radius: 5px;
    }}
    QTabWidget::pane {{
        border: 1px solid {BORDER};
        border-radius: 8px;
        background: {PAGE};
        top: -1px;
    }}
    QTabBar::tab {{
        background: transparent;
        color: {MUTED};
        padding: 7px 18px;
        margin-right: 4px;
        border: 1px solid transparent;
        border-top-left-radius: 8px;
        border-top-right-radius: 8px;
        font-weight: 600;
    }}
    QTabBar::tab:selected {{
        background: {CARD};
        color: {NAVY};
        border: 1px solid {BORDER};
        border-bottom-color: {CARD};
    }}
    QTextEdit#log {{
        background: {CARD};
        border: 1px solid {BORDER};
        border-radius: 8px;
        font-family: "Consolas";
        font-size: 9pt;
        padding: 8px;
    }}
    /* Which set of WIR paths is in use. A plain radio dot was too easy to
       miss, so the selected side is a filled block and the other is a pale
       outline - the state is readable from across the room. Server is navy
       like the checks; Local is teal, because it means "on this PC". */
    QFrame#modeBar {{
        background: #FFFFFF;
        border: 1px solid {BORDER};
        border-radius: 8px;
    }}
    QLabel#modeCaption {{
        color: {MUTED};
        font-size: 9pt;
        font-weight: 600;
    }}
    QLabel#modeState {{
        font-size: 9pt;
        font-weight: 600;
    }}
    QPushButton#modeServer, QPushButton#modeLocal {{
        background: #FFFFFF;
        color: {MUTED};
        border: 2px solid {BORDER};
        border-radius: 6px;
        padding: 7px 26px;
        font-size: 10pt;
        font-weight: 700;
    }}
    QPushButton#modeServer:checked {{
        background: {NAVY};
        border-color: {NAVY};
        color: #FFFFFF;
    }}
    QPushButton#modeLocal:checked {{
        background: {TEAL};
        border-color: {TEAL};
        color: #FFFFFF;
    }}
    QPushButton#modeServer:disabled, QPushButton#modeLocal:disabled {{
        color: #B9BDC5;
        border-color: #ECEEF1;
    }}

    /* Sits opposite the title. Blue to match it, but smaller and italic so
       it still reads as a credit rather than a second heading. */
    QLabel#author {{
        color: {NAVY};
        font-size: 9pt;
        font-style: italic;
        padding-top: 6px;
    }}
    QPushButton#runAll {{
        background: {NAVY};
        color: #FFFFFF;
        border: none;
        border-radius: 8px;
        padding: 11px 20px;
        font-weight: 600;
        font-size: 10.5pt;
    }}
    QPushButton#runAll:hover  {{ background: #2A6395; }}
    QPushButton#runAll:pressed {{ background: #17395A; }}
    QPushButton#runAll:disabled {{ background: #A8B4C0; }}

    /* The last thing you press, and the only one that produces documents.
       Taller and coral against the navy above it, so the eye lands on the
       order of work: check everything, then generate. */
    QPushButton#generate {{
        background: {CORAL};
        color: #FFFFFF;
        border: none;
        border-radius: 8px;
        padding: 16px 20px;
        font-weight: 700;
        font-size: 12pt;
    }}
    QPushButton#generate:hover  {{ background: #E8703F; }}
    QPushButton#generate:pressed {{ background: #B2481F; }}
    QPushButton#generate:disabled {{ background: #D9BDB1; }}
    """


def card_style(accent: str, tint: str, text_colour: str, filled: bool) -> str:
    """A button that looks like a card: accent stripe down the left, flat
    fill, and a real hover state."""
    background = tint if filled else CARD
    hover = tint if not filled else "#F5C4B3"
    return f"""
    QPushButton {{
        background: {background};
        border: 1px solid {BORDER};
        border-left: 4px solid {accent};
        border-radius: 8px;
        padding: 9px 12px;
        text-align: left;
        color: {text_colour};
        font-weight: 600;
        font-size: 10pt;
    }}
    QPushButton:hover {{
        background: {hover};
        border: 1px solid {accent};
        border-left: 4px solid {accent};
    }}
    QPushButton:pressed {{ background: {accent}; color: #FFFFFF; }}
    QPushButton:disabled {{ color: {MUTED}; border-left-color: {BORDER}; }}
    """
