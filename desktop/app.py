"""
Clearance WIR Log Validation — desktop front end.

One workbook holds everything: the WIRs and Clearance sheets, the rules, and
Project-Des. So there is one file to choose, and the four outputs always go to a
fixed folder beside the app. Nothing to configure per run.

The two Import buttons are the Python port of the workbook's own macros: they
pull a raw log off the server or off disk and write it into the master's WIRs
or Clearance sheet.

Work runs on a background thread and every message comes back through a queue.
tkinter is not thread-safe, and a cross-thread widget call fails intermittently
rather than outright, which is the worst way for it to fail.
"""

from __future__ import annotations

import json
import os
import queue
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

APP_ROOT = Path(__file__).resolve().parents[1]
if str(APP_ROOT) not in sys.path:
    sys.path.insert(0, str(APP_ROOT))

from engine.pipeline import PRODUCTS, Level, Product, RunReport, output_dir_for, run  # noqa: E402

APP_TITLE = "Clearance WIR Log Validation"
MASTER_NAME = "WIR_Clearance Log Validation.xlsm"
EXCEL_TYPES = [("Excel workbooks", "*.xlsm *.xlsx *.xls"), ("All files", "*.*")]

COLOURS = {
    Level.INFO: "#1F2933",
    Level.GOOD: "#1B6E3C",
    Level.WARN: "#8A5B06",
    Level.FAIL: "#A32D18",
}


def app_dir() -> Path:
    """Where the app lives — beside the exe once frozen."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return APP_ROOT


def bundled_reference() -> Path:
    """PyInstaller unpacks bundled data to _MEIPASS at runtime."""
    return Path(getattr(sys, "_MEIPASS", APP_ROOT)) / "reference"


def find_master() -> Path | None:
    """Look for the workbook beside the app, then a few folders up.

    A portable deployment keeps the exe and its runtime in their own folder, so
    the workbook usually sits a level or two above rather than alongside.
    """
    here = app_dir()
    for folder in [here, *here.parents][:4]:
        candidate = folder / MASTER_NAME
        if candidate.exists():
            return candidate
    return None


def default_master() -> str:
    found = find_master()
    return str(found) if found else ""


SETTINGS_PATH = app_dir() / "log-validation-settings.json"


def load_settings() -> dict:
    try:
        return json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
    except Exception:  # noqa: BLE001 - settings are a convenience, never required
        return {}


def save_settings(data: dict) -> None:
    try:
        SETTINGS_PATH.write_text(json.dumps(data, indent=2), encoding="utf-8")
    except Exception:  # noqa: BLE001
        pass


class LogValidationApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.queue: queue.Queue = queue.Queue()
        self.worker: threading.Thread | None = None
        self.settings = load_settings()

        root.title(APP_TITLE)
        root.geometry("900x700")
        root.minsize(780, 600)

        self.master_path = tk.StringVar(
            value=self.settings.get("master_path") or default_master()
        )
        self.status = tk.StringVar(value="Ready.")
        self.wanted = {
            product: tk.BooleanVar(
                value=self.settings.get("wanted", {}).get(product.value, True)
            )
            for product in Product
        }

        self._build()
        self.root.after(100, self._drain)
        if not self.master_path.get():
            self._write(f"No {MASTER_NAME} found beside the app — choose it above.", Level.WARN)

    # ------------------------------------------------------------------ build

    def _build(self) -> None:
        outer = ttk.Frame(self.root, padding=14)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(4, weight=1)

        self._source(outer).grid(row=0, column=0, sticky="ew")
        self._imports(outer).grid(row=1, column=0, sticky="ew", pady=(12, 0))
        self._generate(outer).grid(row=2, column=0, sticky="ew", pady=(12, 0))
        self._actions(outer).grid(row=3, column=0, sticky="ew", pady=(12, 0))
        self._log(outer).grid(row=4, column=0, sticky="nsew", pady=(12, 0))

        bar = ttk.Frame(outer)
        bar.grid(row=5, column=0, sticky="ew", pady=(8, 0))
        bar.columnconfigure(0, weight=1)
        ttk.Label(bar, textvariable=self.status, foreground="#4E5A62").grid(
            row=0, column=0, sticky="w"
        )

    def _source(self, parent) -> ttk.LabelFrame:
        frame = ttk.LabelFrame(parent, text="  Log workbook  ", padding=12)
        frame.columnconfigure(0, weight=1)
        ttk.Entry(frame, textvariable=self.master_path).grid(row=0, column=0, sticky="ew")
        ttk.Button(frame, text="Browse…", width=11, command=self._pick_master).grid(
            row=0, column=1, padx=(8, 0)
        )
        ttk.Label(
            frame,
            text="Holds the WIRs and Clearance sheets, the rules, and Project-Des. "
                 "Edit any rule there.",
            foreground="#7D8B95",
            font=("Segoe UI", 8),
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(4, 0))
        return frame

    def _imports(self, parent) -> ttk.LabelFrame:
        frame = ttk.LabelFrame(parent, text="  Import raw logs  ", padding=12)
        buttons = ttk.Frame(frame)
        buttons.grid(row=0, column=0, sticky="w")
        self.import_wir = ttk.Button(
            buttons, text="Import WIR Log", width=20,
            command=lambda: self._start_import("wir"),
        )
        self.import_wir.pack(side="left")
        self.import_clearance = ttk.Button(
            buttons, text="Import Clearance Log", width=22,
            command=lambda: self._start_import("clearance"),
        )
        self.import_clearance.pack(side="left", padx=(8, 0))
        ttk.Label(
            frame,
            text="Replaces the WIRs or Clearance sheet in the workbook above from a "
                 "raw log. Close the workbook in Excel first.",
            foreground="#7D8B95",
            font=("Segoe UI", 8),
        ).grid(row=1, column=0, sticky="w", pady=(6, 0))
        return frame

    def _generate(self, parent) -> ttk.LabelFrame:
        frame = ttk.LabelFrame(parent, text="  Generate  ", padding=12)
        boxes = ttk.Frame(frame)
        boxes.grid(row=0, column=0, sticky="w")
        for index, product in enumerate(Product):
            ttk.Checkbutton(
                boxes, text=PRODUCTS[product].title, variable=self.wanted[product]
            ).grid(row=index // 2, column=index % 2, sticky="w", padx=(0, 32), pady=2)
        self.output_label = ttk.Label(
            frame, text="", foreground="#7D8B95", font=("Segoe UI", 8)
        )
        self.output_label.grid(row=1, column=0, sticky="w", pady=(8, 0))
        self._refresh_output_label()
        self.master_path.trace_add("write", lambda *_: self._refresh_output_label())
        return frame

    def _refresh_output_label(self) -> None:
        self.output_label.configure(text=f"Saves to  {self.output_dir()}")

    def _actions(self, parent) -> ttk.Frame:
        frame = ttk.Frame(parent)
        frame.columnconfigure(1, weight=1)
        self.run_button = ttk.Button(frame, text="Generate", command=self._start_run, width=16)
        self.run_button.grid(row=0, column=0, sticky="w")
        self.progress = ttk.Progressbar(frame, mode="determinate", maximum=100)
        self.progress.grid(row=0, column=1, sticky="ew", padx=12)
        self.open_button = ttk.Button(
            frame, text="Open folder", command=self._open_output, width=14
        )
        self.open_button.grid(row=0, column=2, sticky="e")
        return frame

    def _log(self, parent) -> ttk.LabelFrame:
        frame = ttk.LabelFrame(parent, text="  Progress  ", padding=8)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)
        self.text = tk.Text(
            frame, wrap="word", height=14, font=("Consolas", 9),
            background="#FBFBFC", relief="flat", padx=8, pady=6,
        )
        self.text.grid(row=0, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(frame, command=self.text.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.text.configure(yscrollcommand=scroll.set, state="disabled")
        for level, colour in COLOURS.items():
            self.text.tag_config(level.value, foreground=colour)
        return frame

    # ----------------------------------------------------------------- helpers

    def output_dir(self) -> Path:
        """Beside the workbook, so results land next to the data they describe
        rather than inside the app's own folder."""
        master = self.master_path.get().strip()
        base = Path(master).parent if master else app_dir()
        return output_dir_for(base)

    def _busy(self, busy: bool) -> None:
        state = "disabled" if busy else "normal"
        for widget in (self.run_button, self.import_wir, self.import_clearance):
            widget.configure(state=state)

    def _ready_master(self) -> Path | None:
        path = self.master_path.get().strip()
        if not path:
            messagebox.showwarning(APP_TITLE, "Choose the log workbook first.")
            return None
        if not Path(path).exists():
            messagebox.showwarning(APP_TITLE, f"Not found:\n{path}")
            return None
        return Path(path)

    def _pick_master(self) -> None:
        chosen = filedialog.askopenfilename(
            title="Select the log workbook",
            filetypes=EXCEL_TYPES,
            initialdir=str(app_dir()),
        )
        if chosen:
            self.master_path.set(chosen)

    # ------------------------------------------------------------------ import

    def _start_import(self, which: str) -> None:
        if self.worker and self.worker.is_alive():
            return
        master = self._ready_master()
        if master is None:
            return

        label = "WIR" if which == "wir" else "Clearance"
        source = filedialog.askopenfilename(
            title=f"Select the raw {label} log to import",
            filetypes=EXCEL_TYPES,
            initialdir=self.settings.get("import_dir") or str(app_dir()),
        )
        if not source:
            return
        self.settings["import_dir"] = str(Path(source).parent)

        if not messagebox.askyesno(
            APP_TITLE,
            f"Replace the '{label}' sheet in\n{master.name}\n\nwith the contents of\n"
            f"{Path(source).name}?\n\nThe workbook must be closed in Excel.",
        ):
            return

        self._reset_log()
        self._busy(True)
        self.progress.configure(value=0, maximum=3)
        self.status.set(f"Importing the {label} log…")
        self.worker = threading.Thread(
            target=self._do_import, args=(master, Path(source), which), daemon=True
        )
        self.worker.start()

    def _do_import(self, master: Path, source: Path, which: str) -> None:
        from engine.importer import ImportError_, import_log

        try:
            self.queue.put(("log", f"Reading {source.name}…", Level.INFO))
            result = import_log(master, source, which)
            for skipped in result.skipped:
                self.queue.put(("log", f"  sheet '{skipped}' not in the source — skipped",
                                Level.WARN))
            self.queue.put(("log", f"  {result.headline}", Level.GOOD))
            self.queue.put(("done", None, f"{result.rows} rows imported."))
        except ImportError_ as exc:
            self.queue.put(("log", str(exc), Level.FAIL))
            self.queue.put(("done", None, "Import failed."))
        except Exception as exc:  # noqa: BLE001 - never let the worker die silently
            self.queue.put(("log", f"Unexpected failure: {exc}", Level.FAIL))
            self.queue.put(("done", None, "Import failed."))

    # --------------------------------------------------------------- generate

    def _start_run(self) -> None:
        if self.worker and self.worker.is_alive():
            return
        master = self._ready_master()
        if master is None:
            return

        wanted = [p for p in Product if self.wanted[p].get()]
        if not wanted:
            messagebox.showwarning(APP_TITLE, "Choose at least one workbook to generate.")
            return

        self._remember()
        self._reset_log()
        self._busy(True)
        self.progress.configure(value=0, maximum=max(len(wanted) * 2, 1))
        self.status.set("Working…")
        self.worker = threading.Thread(target=self._do_run, args=(master, wanted), daemon=True)
        self.worker.start()

    def _do_run(self, master: Path, wanted: list[Product]) -> None:
        try:
            report = run(
                source_path=master,
                output_dir=self.output_dir(),
                config_dir=bundled_reference(),
                wanted=wanted,
                progress=lambda message, level: self.queue.put(("log", message, level)),
            )
            self.queue.put(("done", report, ""))
        except Exception as exc:  # noqa: BLE001
            self.queue.put(("log", f"Unexpected failure: {exc}", Level.FAIL))
            self.queue.put(("done", None, "Stopped."))

    # ------------------------------------------------------------------ events

    def _drain(self) -> None:
        try:
            while True:
                kind, *payload = self.queue.get_nowait()
                if kind == "log":
                    message, level = payload
                    self._write(message, level)
                    if not message.startswith(" "):
                        self.progress.step(1)
                elif kind == "done":
                    self._finish(payload[0], payload[1])
        except queue.Empty:
            pass
        self.root.after(100, self._drain)

    def _reset_log(self) -> None:
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        self.text.configure(state="disabled")

    def _write(self, message: str, level: Level = Level.INFO) -> None:
        self.text.configure(state="normal")
        self.text.insert("end", message + "\n", level.value)
        self.text.see("end")
        self.text.configure(state="disabled")

    def _finish(self, report: RunReport | None, note: str) -> None:
        self._busy(False)
        self.progress.configure(value=self.progress["maximum"])

        if report is None:
            self.status.set(note or "Done.")
            return

        done = [r for r in report.results if r.ok]
        failed = [r for r in report.results if not r.ok]
        self._write("", Level.INFO)
        if done:
            self._write(f"Wrote {len(done)} workbook(s) to {self.output_dir()}", Level.GOOD)
        for result in failed:
            self._write(f"{result.spec.title}: {result.error}", Level.WARN)

        if failed and done:
            self.status.set(f"{len(done)} generated, {len(failed)} failed.")
        elif failed or report.failed_early:
            self.status.set("Nothing generated — see the messages above.")
        else:
            self.status.set("Done.")

    def _open_output(self) -> None:
        target = self.output_dir()
        target.mkdir(parents=True, exist_ok=True)
        os.startfile(target)  # noqa: S606 - Windows shell open on our own folder

    def _remember(self) -> None:
        self.settings.update(
            {
                "master_path": self.master_path.get(),
                "wanted": {p.value: self.wanted[p].get() for p in Product},
            }
        )
        save_settings(self.settings)


def main() -> int:
    if "--selftest" in sys.argv:
        from desktop.selftest import main as selftest

        return selftest()

    try:
        from ctypes import windll

        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:  # noqa: BLE001 - not Windows, or an older build
        pass

    root = tk.Tk()
    try:
        ttk.Style().theme_use("vista")
    except tk.TclError:
        pass
    LogValidationApp(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
