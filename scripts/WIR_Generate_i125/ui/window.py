"""
The window.

Two tabs: the actions, and the log of what the last run did. Every task
runs on a worker thread so the window never freezes while Excel is busy.

Anything that needs a file or folder from the user is asked for here, on
the GUI thread, and handed to the worker - a dialog opened from a worker
thread is a crash waiting to happen.
"""

from __future__ import annotations

import os

from PySide6.QtCore import QObject, QThread, Qt, Signal
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QApplication, QButtonGroup, QFileDialog, QFrame, QHBoxLayout, QLabel,
    QMainWindow, QMessageBox, QProgressDialog, QPushButton, QTabWidget,
    QTextEdit, QVBoxLayout, QWidget,
)

from core import config, pathmap, reference, tasks
from core.findings import RunResult
from core.progress import Progress
from core.workbook import LiveWorkbook, open_state

from . import theme

APP_NAME = "WIR Generate Tools"
AUTHOR = "Author: Mina Kamel"

# section | title | description | task id
ACTIONS = [
    ("tools", "Import Previous WIR Paths", "Fill column AA with approved previous WIRs", "import_prev"),
    ("tools", "Extract PDFs", "Copy files from the visible rows to a folder", "extract_pdfs"),
    ("tools", "Extract WIR Cover Page", "First page of each commented WIR, share to local", "extract_covers"),
    ("tools", "Clear Any Sheet", "Wipe rows 3-1000 on the active sheet", "clear_any"),
    ("verify", "Fix Attachment Links", "Repair drive letters, quotes, missing Common", "fix_attachments"),
    ("verify", "Fix WIR Path Prefix", "Register paths to the local extract folder", "fix_prefix"),
    ("verify", "Check Areas", "Suggest the correct Area text into column K", "check_areas"),
    ("verify", "Check WIRs Before Generate", "Duplicates, precedence, areas and links", "preflight"),
    ("generate", "Import WIR Log", "Reload the WIRs register from the log file", "import_data"),
    ("generate", "Copy To History", "Append the visible rows to WIR_History", "copy_history"),
    ("generate", "Clear WIR Sheet", "Wipe rows 3-100, restore template formulas", "clear_wir"),
]

# Generate WIRs is not in a column. It is the last thing you press and the
# only one that produces documents, so it sits on its own beneath the
# checks - the two wide buttons read as the order of work.

SECTION_TITLES = {
    "tools": "Tools",
    "verify": "Verify and maintain",
    "generate": "Generate",
}

# Tasks that destroy data get a confirmation first.
CONFIRM = {
    "clear_any": "Clear rows 3-1000 on the active sheet?\n\nThis cannot be undone.",
    "clear_wir": "Clear the WIR log and restore the template row?\n\nThis cannot be undone.",
}


# Tasks that report item-by-item and can be stopped part way. Everything
# else gets a busy bar, because a fake percentage is worse than none.
CANCELLABLE = {"extract_covers", "generate"}


class Worker(QObject):
    finished = Signal(list)
    failed = Signal(str)
    progressed = Signal(int, int, str)

    def __init__(self, task: str, params: dict):
        super().__init__()
        self.task = task
        self.params = params
        # Emitting from this thread to the window is safe - Qt queues the
        # signal onto the GUI thread.
        self.progress = Progress(
            lambda done, total, message: self.progressed.emit(done, total, message)
        )

    def run(self) -> None:
        try:
            # A standalone task must not attach: it has nothing to do with
            # the workbook, and demanding Excel be open to run it would be
            # an invented requirement.
            if self.task in tasks.NO_WORKBOOK:
                self.finished.emit(self._dispatch_standalone())
                return

            book = LiveWorkbook.attach()
            with book:
                self.finished.emit(self._dispatch(book))
        except Exception as exc:                   # noqa: BLE001 - shown to the user
            self.failed.emit(str(exc))

    def _dispatch_standalone(self) -> list[RunResult]:
        if self.task == "extract_covers":
            return [tasks.extract_covers.run(progress=self.progress)]
        raise ValueError(f"Unknown standalone task '{self.task}'.")

    def _dispatch(self, book: LiveWorkbook) -> list[RunResult]:
        task = self.task

        if task == "clear_any":
            return [tasks.clear_sheets.clear_any(book)]
        if task == "clear_wir":
            return [tasks.clear_sheets.clear_log(book)]
        if task == "fix_attachments":
            return [tasks.fix_attachments.run(book)]
        if task == "fix_prefix":
            return [tasks.fix_prefix.run(book)]
        if task == "extract_pdfs":
            return [tasks.extract_pdfs.run(book, self.params.get("destination", ""))]
        if task == "import_data":
            return [tasks.import_register.run(
                book, self.params.get("source", ""),
                self.params.get("source_action", "discard"),
            )]
        if task == "copy_history":
            return [tasks.history.run(book, self.params.get("email", False),
                                      self.params.get("greeting", ""))]

        ref = reference.load(book)
        if task == "generate":
            # Import Previous WIR Paths runs first - see tasks.generate_wirs.
            return tasks.generate_wirs(book, ref, self.params.get("output", ""),
                                       progress=self.progress)
        if task == "check_areas":
            return [tasks.check_areas.run(book, ref)]
        if task == "preflight":
            return [tasks.preflight.run(book, ref)]
        if task == "import_prev":
            return [tasks.import_previous.run(book, ref)]
        if task == "run_all":
            return tasks.run_all(book, ref)

        raise ValueError(f"Unknown task '{task}'.")


class ActionCard(QPushButton):
    def __init__(self, title: str, description: str, section: str, filled: bool = False):
        super().__init__()
        accent, tint, text_colour = theme.SECTIONS[section]
        self.setStyleSheet(theme.card_style(accent, tint, text_colour, filled))
        self.setCursor(Qt.PointingHandCursor)
        self.setMinimumHeight(54)
        self.setToolTip(description)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 3, 8, 3)
        layout.setSpacing(1)

        name = QLabel(title)
        name.setStyleSheet(
            f"color: {text_colour}; font-weight: 600; background: transparent;"
        )
        note = QLabel(description)
        note.setWordWrap(True)
        note.setStyleSheet(
            f"color: {theme.MUTED}; font-size: 8pt; background: transparent;"
        )
        layout.addWidget(name)
        layout.addWidget(note)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.setMinimumSize(800, 520)
        self.setStyleSheet(theme.stylesheet())
        self._thread: QThread | None = None
        self._worker: Worker | None = None
        self._progress_dialog: QProgressDialog | None = None
        self._buttons: list[QPushButton] = []

        root = QWidget()
        self.setCentralWidget(root)
        outer = QVBoxLayout(root)
        outer.setContentsMargins(18, 14, 18, 12)
        outer.setSpacing(10)

        # Title left, author right, on one line.
        header = QHBoxLayout()
        title = QLabel(APP_NAME)
        title.setObjectName("title")
        author = QLabel(AUTHOR)
        author.setObjectName("author")
        header.addWidget(title)
        header.addStretch(1)
        header.addWidget(author, 0, Qt.AlignRight | Qt.AlignVCenter)
        outer.addLayout(header)

        self.tabs = QTabWidget()
        self.tabs.addTab(self._actions_tab(), "Actions")
        self.tabs.addTab(self._log_tab(), "Last run")
        outer.addWidget(self.tabs)

        self.status = QLabel("Ready. Open the WIR workbook in Excel, then pick an action.")
        self.status.setObjectName("status")
        outer.addWidget(self.status)

        # Server every time the program starts, whatever the sheet was
        # left on. Quietly - no workbook open is an ordinary state at
        # this point, and the window has nothing to say about it yet.
        self._reset_mode()

    # -- tabs ---------------------------------------------------------------

    def _actions_tab(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(12)

        columns = QHBoxLayout()
        columns.setSpacing(14)
        for key in ("tools", "verify", "generate"):
            columns.addWidget(self._section(key))
        layout.addLayout(columns)

        run_all = QPushButton("Run all checks    Fix Attachment Links  +  Check Areas  +  Check WIRs")
        run_all.setObjectName("runAll")
        run_all.setCursor(Qt.PointingHandCursor)
        run_all.setToolTip(
            "Runs the three verify steps in order: repairs the links, tidies the "
            "areas, then checks every pending row."
        )
        run_all.clicked.connect(lambda: self._start("run_all"))
        self._buttons.append(run_all)
        layout.addWidget(run_all)

        # Server or Local. A plain radio dot was too easy to miss, so
        # this is a two-state block: the side in use is filled, the other
        # is a pale outline, and the caption spells it out in words.
        modes = QFrame()
        modes.setObjectName("modeBar")
        mode_row = QHBoxLayout(modes)
        mode_row.setContentsMargins(14, 8, 14, 8)
        mode_row.setSpacing(10)

        caption = QLabel("WIR PATHS")
        caption.setObjectName("modeCaption")
        mode_row.addWidget(caption)

        self.server_button = QPushButton("SERVER")
        self.server_button.setObjectName("modeServer")
        self.local_button = QPushButton("LOCAL")
        self.local_button.setObjectName("modeLocal")
        for button in (self.server_button, self.local_button):
            button.setCheckable(True)
            button.setCursor(Qt.PointingHandCursor)
        self.server_button.setChecked(True)
        self.server_button.setToolTip(
            "Previous-activity paths come from the register's column N, "
            "as filed on the share."
        )
        self.local_button.setToolTip(
            "Previous-activity paths come from the register's column AD, "
            "the copy on this PC."
        )

        self.mode_group = QButtonGroup(self)
        self.mode_group.setExclusive(True)
        self.mode_group.addButton(self.server_button, pathmap.SERVER)
        self.mode_group.addButton(self.local_button, pathmap.LOCAL)
        self.mode_group.idClicked.connect(self._mode_changed)

        mode_row.addWidget(self.server_button)
        mode_row.addWidget(self.local_button)

        self.mode_state = QLabel()
        self.mode_state.setObjectName("modeState")
        mode_row.addWidget(self.mode_state)
        mode_row.addStretch(1)
        layout.addWidget(modes)

        generate = QPushButton("Generate WIRs")
        generate.setObjectName("generate")
        generate.setCursor(Qt.PointingHandCursor)
        generate.setToolTip(
            "Refreshes column AA (Previous Activity) first, then builds "
            "cover, checklist and attachments for every pending row and "
            "merges each into one PDF. Marks the rows Done."
        )
        generate.clicked.connect(lambda: self._start("generate"))
        self._buttons.append(generate)
        layout.addWidget(generate)

        return page

    def _section(self, key: str) -> QWidget:
        accent, _, _ = theme.SECTIONS[key]
        frame = QFrame()
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(7)

        header = QLabel(SECTION_TITLES[key])
        header.setStyleSheet(
            f"background: {accent}; color: #FFFFFF; font-weight: 600; "
            "font-size: 8.5pt; padding: 5px 10px; border-radius: 5px;"
        )
        layout.addWidget(header)

        for section, title, description, task in ACTIONS:
            if section != key:
                continue
            card = ActionCard(title, description, section)
            card.clicked.connect(lambda _=False, t=task: self._start(t))
            self._buttons.append(card)
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

        question = CONFIRM.get(task)
        if question:
            answer = QMessageBox.question(
                self, APP_NAME, question,
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No,
            )
            if answer != QMessageBox.Yes:
                return

        params = self._ask_for_paths(task)
        if params is None:
            return

        self._set_busy(True, task)
        self._thread = QThread(self)
        self._worker = Worker(task, params)
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.progressed.connect(self._on_progress)
        self._worker.finished.connect(self._done)
        self._worker.failed.connect(self._error)
        self._worker.finished.connect(self._thread.quit)
        self._worker.failed.connect(self._thread.quit)
        self._thread.finished.connect(self._worker.deleteLater)
        self._show_progress(task)
        self._thread.start()

    def _show_mode(self, mode: int, note: str = "") -> None:
        """Keep the caption saying, in words, what the colours show."""
        if mode == pathmap.LOCAL:
            text, colour = "using column AD on this PC", theme.TEAL
        else:
            text, colour = "using column N on the server", theme.NAVY
        self.mode_state.setText(f"{text}{note}")
        self.mode_state.setStyleSheet(f"color: {colour};")

    def _reset_mode(self) -> None:
        """Server every time the program starts, whatever the sheet held."""
        self.server_button.setChecked(True)
        landed = False
        try:
            book = LiveWorkbook.attach()
            with book:
                landed = pathmap.write_mode(book, pathmap.SERVER)
        except Exception:                          # noqa: BLE001
            pass
        self._show_mode(pathmap.SERVER, "" if landed else "  (workbook not open)")

    def _mode_changed(self, mode: int) -> None:
        """Write the choice into Main!Y1 so the sheet formulas follow it.

        Done here rather than when a task runs, because the HYPERLINK
        formulas in the log recalculate the moment the cell changes - the
        user sees the links repoint as they click.
        """
        name = "LOCAL" if mode == pathmap.LOCAL else "SERVER"
        landed = False
        try:
            book = LiveWorkbook.attach()
            with book:
                landed = pathmap.write_mode(book, mode)
        except Exception:                          # noqa: BLE001
            pass                                   # no workbook open is ordinary

        cell = f"{config.MODE_SHEET}!Y{config.MODE_ROW}"
        if landed:
            self._show_mode(mode)
            self.status.setText(f"WIR paths set to {name} - {cell} = {mode}.")
        else:
            self._show_mode(mode, "  (workbook not open)")
            self.status.setText(
                f"WIR paths set to {name}, but {cell} could not be written. "
                "The sheet formulas will not follow until the workbook is open."
            )

    # -- progress -----------------------------------------------------------

    def _show_progress(self, task: str) -> None:
        title = task.replace("_", " ").title()
        cancellable = task in CANCELLABLE

        # 0..0 is Qt's busy bar - the right thing to show while a task
        # that cannot count its work is running.
        dialog = QProgressDialog(f"{title}...", "Stop" if cancellable else "", 0, 0, self)
        dialog.setWindowTitle(APP_NAME)
        dialog.setWindowModality(Qt.WindowModal)
        dialog.setMinimumWidth(460)
        dialog.setMinimumDuration(0)        # show at once, not after 4 seconds
        dialog.setAutoClose(False)
        dialog.setAutoReset(False)
        dialog.setValue(0)

        if cancellable:
            dialog.canceled.connect(self._cancel)
        else:
            dialog.setCancelButton(None)

        self._progress_dialog = dialog
        dialog.show()

    def _on_progress(self, done: int, total: int, message: str) -> None:
        dialog = self._progress_dialog
        if dialog is None:
            return
        if total > 0:
            dialog.setMaximum(total)
            dialog.setValue(done)
            dialog.setLabelText(f"{done} of {total}\n{message}")
        else:
            dialog.setLabelText(message or "Working...")

    def _cancel(self) -> None:
        """Asks the task to stop; it finishes the file it is on and
        returns what it has done. The dialog stays up until it does."""
        if self._worker is not None:
            self._worker.progress.cancel()
        if self._progress_dialog is not None:
            self._progress_dialog.setLabelText("Stopping...")

    def _close_progress(self) -> None:
        if self._progress_dialog is not None:
            self._progress_dialog.close()
            self._progress_dialog = None

    def _ask_for_paths(self, task: str) -> dict | None:
        """Dialogs belong on this thread, never inside the worker."""
        if task == "import_data":
            path, _ = QFileDialog.getOpenFileName(
                self, "Select the WIR register", "",
                "Excel files (*.xlsm *.xlsx *.xls);;All files (*.*)",
            )
            if not path:
                return None
            path = os.path.normpath(path)

            # Importing closes the register afterwards. If the user
            # already had it open, closing it decides the fate of anything
            # they had not saved - so that is their call, not ours.
            is_open, unsaved = open_state(path)
            if not is_open:
                return {"source": path, "source_action": "discard"}

            box = QMessageBox(self)
            box.setWindowTitle(APP_NAME)
            box.setIcon(QMessageBox.Question)
            box.setText(f"{os.path.basename(path)} is open in Excel.")
            warning = "It has unsaved changes.\n\n" if unsaved else ""
            box.setInformativeText(
                warning + "Importing will close it. What should happen to it?"
            )
            save = box.addButton("Save and close", QMessageBox.AcceptRole)
            discard = box.addButton("Close without saving", QMessageBox.DestructiveRole)
            box.addButton("Cancel", QMessageBox.RejectRole)
            box.setDefaultButton(save if unsaved else discard)
            box.exec()

            clicked = box.clickedButton()
            if clicked is save:
                return {"source": path, "source_action": "save"}
            if clicked is discard:
                return {"source": path, "source_action": "discard"}
            return None

        if task == "extract_pdfs":
            folder = QFileDialog.getExistingDirectory(self, "Copy the files into")
            return None if not folder else {"destination": os.path.normpath(folder)}

        # Generate is not asked where to put things: it is always the WIRs
        # folder beside the workbook.

        if task == "copy_history":
            # Asked here rather than left as a checkbox on the tab: a
            # setting sitting off to one side is easy to miss, and the
            # answer is only wanted at the moment the button is pressed.
            answer = QMessageBox.question(
                self, "Copy To History",
                "The visible rows will be copied to WIR_History.\n\n"
                "Draft the submission email in Outlook as well?\n\n"
                "The draft opens for you to check - it is never sent.",
                QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel,
                QMessageBox.Yes,
            )
            if answer == QMessageBox.Cancel:
                return None
            return {"email": answer == QMessageBox.Yes, "greeting": ""}

        return {}

    def _set_busy(self, busy: bool, task: str = "") -> None:
        for button in self._buttons:
            button.setEnabled(not busy)
        if busy:
            self.status.setText(f"Running {task.replace('_', ' ')}...")

    def _done(self, results: list[RunResult]) -> None:
        lines: list[str] = []
        errors = warnings = 0
        for result in results:
            lines.append(f"=== {result.title} ===")
            if not result.ok:
                lines.append(f"  {result.headline}")
            lines.extend(f"  {detail}" for detail in result.details)
            lines.append("")
            errors += result.errors
            warnings += result.warnings

        self.log.setPlainText("\n".join(lines).rstrip())
        self.tabs.setCurrentIndex(1)

        if not results[-1].ok:
            self.status.setText(results[-1].headline)
        elif errors:
            self.status.setText(f"Done - {errors} error(s), {warnings} warning(s).")
        elif warnings:
            self.status.setText(f"Done - {warnings} warning(s), no errors.")
        else:
            self.status.setText("Done.")

        self._finish()

    def _error(self, message: str) -> None:
        self.log.setPlainText(message)
        self.tabs.setCurrentIndex(1)
        self.status.setText(message.split("\n")[0])
        self._finish()

    def _finish(self) -> None:
        self._close_progress()
        self._set_busy(False)
        self._thread = None
        self._worker = None


def launch() -> int:
    app = QApplication.instance() or QApplication([])
    window = MainWindow()
    window.show()
    return app.exec()
