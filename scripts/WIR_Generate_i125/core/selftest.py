"""
Proves the packaged program is complete.

    "WIR Generate Tools.exe" --selftest

Every failure this catches would otherwise reach a user's machine as a
"No module named ..." halfway through a run. The classic one is
win32timezone: pywin32 does not import it until COM first hands back a
date, which is long after PyInstaller has finished reading the imports,
so a build that looks fine fails the first time it touches a dated cell.

That is why the check does not only import things. It makes COM produce a
date, starts a Qt application, and merges a real PDF - those are the three
moments a missing piece actually shows itself.

The report is written next to the executable as well as printed, because
the release build is --noconsole and has nowhere to print to.
"""

from __future__ import annotations

import importlib
import io
import os
import sys
import traceback

# Every module of the program. Named rather than discovered, so a module
# that gets left out of the bundle is a failure and not just an absence.
OWN_MODULES = [
    "core.config",
    "core.findings",
    "core.paths",
    "core.pdf",
    "core.reference",
    "core.report",
    "core.text",
    "core.workbook",
    "core.tasks",
    "core.tasks.check_areas",
    "core.tasks.clear_sheets",
    "core.tasks.extract_pdfs",
    "core.tasks.fix_attachments",
    "core.tasks.fix_prefix",
    "core.tasks.generate",
    "core.tasks.history",
    "core.tasks.import_previous",
    "core.tasks.import_register",
    "core.tasks.preflight",
    "ui.theme",
    "ui.window",
]

# Third-party modules the program reaches for, including the ones it only
# reaches for at the worst possible moment.
THIRD_PARTY = [
    "pythoncom",
    "pywintypes",
    "win32api",
    "win32com",
    "win32com.client",
    "win32com.client.dynamic",
    "win32timezone",          # the one that went missing
    "pypdf",
    "PySide6.QtCore",
    "PySide6.QtGui",
    "PySide6.QtWidgets",
]


def _check_imports(names: list[str]) -> list[str]:
    problems = []
    for name in names:
        try:
            importlib.import_module(name)
        except Exception as exc:                       # noqa: BLE001
            problems.append(f"{name}: {exc}")
    return problems


def _check_com_date() -> list[str]:
    """The real reproduction of the win32timezone failure.

    FileSystemObject is on every Windows machine and DateLastModified
    comes back as a COM date, so this walks the exact code path that used
    to blow up on a dated cell - without needing Excel to be open.
    """
    import win32com.client

    fso = win32com.client.Dispatch("Scripting.FileSystemObject")
    stamp = fso.GetFile(sys.executable).DateLastModified
    return [] if stamp is not None else ["COM returned no date"]


def _check_qt() -> list[str]:
    """Starts Qt for real. A bundle can hold every PySide6 module and
    still fail here with "could not find the Qt platform plugin", because
    the platform plugins are DLLs rather than imports.
    """
    from PySide6.QtWidgets import QApplication

    app = QApplication.instance() or QApplication([])
    name = app.platformName()
    return [] if name else ["Qt started with no platform plugin"]


def _check_pdf_merge() -> list[str]:
    """Builds two one-page PDFs and merges them, which is what Generate
    does to every WIR."""
    from pypdf import PdfReader, PdfWriter

    def one_page() -> bytes:
        writer = PdfWriter()
        writer.add_blank_page(width=200, height=200)
        buffer = io.BytesIO()
        writer.write(buffer)
        writer.close()
        return buffer.getvalue()

    merged = PdfWriter()
    for _ in range(2):
        merged.append(PdfReader(io.BytesIO(one_page())))
    out = io.BytesIO()
    merged.write(out)
    merged.close()

    pages = len(PdfReader(io.BytesIO(out.getvalue())).pages)
    return [] if pages == 2 else [f"merged to {pages} pages, expected 2"]


CHECKS = [
    ("program modules", lambda: _check_imports(OWN_MODULES)),
    ("third-party modules", lambda: _check_imports(THIRD_PARTY)),
    ("COM dates (win32timezone)", _check_com_date),
    ("Qt platform plugin", _check_qt),
    ("PDF merge", _check_pdf_merge),
]


def _report_path() -> str:
    folder = (
        os.path.dirname(sys.executable) if getattr(sys, "frozen", False)
        else os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    )
    return os.path.join(folder, "selftest.log")


def run() -> int:
    lines = [
        "WIR Generate Tools self-test",
        f"  running from : {sys.executable}",
        f"  frozen       : {'yes' if getattr(sys, 'frozen', False) else 'no'}",
        f"  python       : {sys.version.split()[0]}",
        "",
    ]

    failed = 0
    for title, check in CHECKS:
        try:
            problems = check()
        except Exception:                              # noqa: BLE001
            problems = traceback.format_exc().strip().splitlines()[-3:]

        if problems:
            failed += 1
            lines.append(f"FAIL  {title}")
            lines.extend(f"        {problem}" for problem in problems)
        else:
            lines.append(f"ok    {title}")

    lines.append("")
    lines.append(
        "Everything the program needs is present."
        if not failed else
        f"{failed} check(s) failed - this build is not safe to hand out."
    )

    report = "\n".join(lines)
    print(report)
    try:
        with open(_report_path(), "w", encoding="utf-8") as handle:
            handle.write(report + "\n")
    except OSError:
        pass                                           # a read-only folder is not a failure

    return 1 if failed else 0
