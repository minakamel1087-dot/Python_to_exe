"""
The window.

Two tabs: the actions, and the log of what the last run did. Every task
runs on a worker thread so the window never freezes while Excel is busy -
which matters, because pre-flight against a full register takes a couple
of seconds.
"""

from __future__ import annotations

from PySide6.QtCore import QObject, QThread, Qt, Signal
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QApplication, QFrame, QGridLayout, QHBoxLayout, QLabel, QMainWindow,
    QPushButton, QTabWidget, QTextEdit, QVBoxLayout, QWidget,
)

from core import reference, tasks
from core.findings import RunResult
from core.workbook import LiveWorkbook
from . import theme

# section key | title | description | task name
ACTIONS = [
    ("tools", "Import WIR Data", "Reload the WIRs register from the log file", "import_data"),
    ("tools", "Import Previous WIR Paths", "Fill column AA with approved previous WIRs", "import_prev"),
    ("tools", "Extract PDFs", "Copy files from the visible rows to a folder", "extract_pdfs"),
    ("tools", "Clear Any Sheet", "Wipe rows 3-1000 on the active sheet", "clear_any"),
    ("verify", "Fix Links", "Attachment paths and the register prefix", "fix_links"),
    ("verify", "Check Areas", "Suggest the correct Area text into column K", "check_areas"),
    ("verify", "Check WIRs Before Generate", "Duplicates, precedence, areas and links", "preflight"),
    ("generate", "Generate WIRs", "Create covers, checklists and attachments", "generate"),
    ("generate", "Copy To History", "Append the visible rows to WIR_History", "copy_history"),
    ("generate", "Clear WIR Sheet", "Wipe rows 3-100, restore template formulas", "clear_wir"),
]

SECTION_TITLES = {
    "tools": "Tools",
    "verify": "Verify and maintain",
    "generate": "Generate",
}

# Tasks that still live in the workbook's VBA. Running them here calls the
# macro rather than reimplementing it - the sheet rendering and the Outlook
# mail are Excel's job either way.
MACRO_TASKS = {
    "import_data": "ImportWIRData",
    "import_prev": "ImportPreviousWIRPaths",
    "extract_pdfs": "CopyFilesFromVisibleRows_ELE",
    "clear_any": "Clear_Any_SHEET",
    "generate": "ELE_WIR_Generate",
    "copy_history": "CopyToHistory",
    "clear_wir": "Clear_WIR_Generate_SHEET",
}


class Worker(QObject):
    """Runs one task against the live workbook, off the UI thread."""

    finished = Signal(list)
    failed = Signal(str)

    def __init__(self, task: str):
        super().__init__()
        self.task = task

    def run(self) -> None:
        try:
            book = LiveWorkbook.attach()
            with book:
                if self.task in MACRO_TASKS:
                    book.excel.Run(MACRO_TASKS[self.task])
                    result = RunResult(MACRO_TASKS[self.task])
                    result.line("Ran the workbook macro.")
                    self.finished.emit([result])
                    return

                ref = reference.load(book)
                if self.task == "fix_links":
                    self.finished.emit([tasks.fix_links.run(book)])
                elif self.task == "check_areas":
                    self.finished.emit([tasks.check_areas.run(book, ref)])
                elif self.task == "preflight":
                    self.finished.emit([tasks.preflight.run(book, ref)])
                elif self.task == "run_all":
                    self.finished.emit(tasks.run_all(book, ref))
                else:
                    self.failed.emit(f"Unknown task '{self.task}'.")
        except Exception as exc:                      # noqa: BLE001 - shown to the user
            self.failed.emit(str(exc))


class ActionCard(QPushButton):
    def __init__(self, title: str, description: str, section: str, filled: bool = False):
        super().__init__()
        accent, tint, text_colour = theme.SECTIONS[section]
        self.setStyleSheet(theme.card_style(accent, tint, text_colour, filled))
        self.setCursor(Qt.PointingHandCursor)
        self.setMinimumHeight(52)
        self.setToolTip(description)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(6, 2, 6, 2)
        layout.setSpacing(0)

        name = QLabel(title)
        name.setStyleSheet(f"color: {text_colour}; font-weight: 600; background: transparent;")
        note = QLabel(description)
        note.setStyleSheet(f"color: {theme.MUTED}; font-size: 8pt; background: transparent;")
        layout.addWidget(name)
        layout.addWidget(note)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("WIR Tools")
        self.setMinimumSize(760, 470)
        self.setStyleSheet(theme.stylesheet())
        self._thread: QThread | None = None
        self._cards: list[QPushButton] = []

        root = QWidget()
        self.setCentralWidget(root)
        outer = QVBoxLayout(root)
        outer.setContentsMargins(18, 14, 18, 14)
        outer.setSpacing(10)

        title = QLabel("WIR Tools")
        title.setObjectName("title")
        outer.addWidget(title)

        self.tabs = QTabWidget()
        self.tabs.addTab(self._actions_tab(), "Actions")
        self.tabs.addTab(self._log_tab(), "Last run")
        outer.addWidget(self.tabs)

        self.status = QLabel("Ready. The workbook must be open in Excel.")
        self.status.setObjectName("status")
        outer.addWidget(self.status)

    # -- tabs ---------------------------------------------------------------

    def _actions_tab(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(12)

        columns = QHBoxLayout()
        columns.setSpacing(14)
        for section in ("tools", "verify", "generate"):
            columns.addWidget(self._section(section))
        layout.addLayout(columns)

        run_all = QPushButton("Run all checks   -   Fix Links, Check Areas, Check WIRs")
        run_all.setObjectName("runAll")
        run_all.setCursor(Qt.PointingHandCursor)
        run_all.setToolTip(
            "Runs the three verify steps in order: repairs the links, tidies the "
            "areas, then checks every pending row."
        )
        run_all.clicked.connect(lambda: self._start("run_all"))
        self._cards.append(run_all)
        layout.addWidget(run_all)

        return page

    def _section(self, key: str) -> QWidget:
        accent, _, _ = theme.SECTIONS[key]
        frame = QFrame()
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(7)

        header = QLabel(SECTION_TITLES[key])
        header.setObjectName("sectionHeader")
        header.setStyleSheet(
            f"background: {accent}; color: #FFFFFF; font-weight: 600; "
            "font-size: 8.5pt; padding: 5px 10px; border-radius: 5px;"
        )
        layout.addWidget(header)

        for section, title, description, task in ACTIONS:
            if section != key:
                continue
            filled = task == "generate"
            card = ActionCard(title, description, section, filled=filled)
            card.clicked.connect(lambda _=False, t=task: self._start(t))
            self._cards.append(card)
            layout.addWidget(card)

        layout.addStretch(1)
        return frame

    def _log_tab(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(14, 14, 14, 14)

        self.log = QTextEdit()
        self.log.setObjectName("log")
        self.log.setReadOnly(True)
        self.log.setFont(QFont("Consolas", 9))
        self.log.setPlainText("Nothing has run yet.")
        layout.addWidget(self.log)
        return page

    # -- running ------------------------------------------------------------

    def _start(self, task: str) -> None:
        if self._thread is not None:
            return

        self._set_busy(True, task)
        self._thread = QThread(self)
        worker = Worker(task)
        worker.moveToThread(self._thread)
        self._thread.started.connect(worker.run)
        worker.finished.connect(self._done)
        worker.failed.connect(self._error)
        worker.finished.connect(self._thread.quit)
        worker.failed.connect(self._thread.quit)
        self._thread.finished.connect(worker.deleteLater)
        self._worker = worker
        self._thread.start()

    def _set_busy(self, busy: bool, task: str = "") -> None:
        for card in self._cards:
            card.setEnabled(not busy)
        if busy:
            self.status.setText(f"Running {task.replace('_', ' ')}...")

    def _done(self, results: list[RunResult]) -> None:
        lines: list[str] = []
        errors = warnings = 0
        for result in results:
            lines.append(f"=== {result.title} ===")
            if not result.ok:
                lines.append(f"  {result.headline}")
            lines.extend(f"  {d}" for d in result.details)
            lines.append("")
            errors += result.errors
            warnings += result.warnings

        self.log.setPlainText("\n".join(lines).rstrip())
        self.tabs.setCurrentIndex(1)

        headline = results[-1].headline if not results[-1].ok else ""
        if headline:
            self.status.setText(headline)
        elif errors:
            self.status.setText(f"Done - {errors} error(s), {warnings} warning(s).")
        elif warnings:
            self.status.setText(f"Done - {warnings} warning(s), no errors.")
        else:
            self.status.setText("Done - nothing to report.")

        self._finish()

    def _error(self, message: str) -> None:
        self.log.setPlainText(message)
        self.tabs.setCurrentIndex(1)
        self.status.setText(message.split("\n")[0])
        self._finish()

    def _finish(self) -> None:
        self._set_busy(False)
        self._thread = None


def launch() -> int:
    app = QApplication.instance() or QApplication([])
    window = MainWindow()
    window.show()
    return app.exec()
