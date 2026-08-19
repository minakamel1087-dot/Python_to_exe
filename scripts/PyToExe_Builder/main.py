"""
PyToExe Builder — the GitHub workflow's job, done on this machine.

Point it at a project folder and it produces the .exe, using the same rules as
build-exe.ps1: a .spec wins over main.py, requirements.txt is found by walking
up from the entry point, pipreqs fills in when there isn't one, and build.args
carries per-project PyInstaller flags.

It needs Python installed on this machine. PyInstaller compiles a program by
importing it and tracing what it pulls in, which takes a real interpreter with
a complete standard library — a frozen exe's own runtime is stripped and can't
do it. So this tool finds Python, keeps its own private virtualenv out of the
way in LOCALAPPDATA, and drives PyInstaller as a subprocess.
"""

import os
import queue
import shutil
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, ttk

APP_NAME = "PyToExe Builder"
HOME = Path(os.environ.get("LOCALAPPDATA", Path.home())) / "PyToExe"
VENV = HOME / "venv"
WORKPATH = HOME / "build"
PYINSTALLER_PIN = "pyinstaller==6.11.1"

# Sentinels the worker thread sends through the log queue to report the outcome.
DONE = "\x00done"
FAILED = "\x00failed"

# Keeps a console window from flashing up for every subprocess — this is a
# windowed build, so those flashes are the only thing the user would see.
NO_WINDOW = subprocess.CREATE_NO_WINDOW if hasattr(subprocess, "CREATE_NO_WINDOW") else 0


# ---------------------------------------------------------------------------
# Finding things
# ---------------------------------------------------------------------------


def can_build(python: Path) -> bool:
    """Whether this interpreter can host a build.

    Embeddable Python distributions — the portable kind an app ships to run on
    a machine with nothing installed — have venv and ensurepip stripped out.
    They run programs perfectly well and cannot build them at all, so they have
    to be filtered out here rather than failing later with something obscure.
    """
    try:
        probe = subprocess.run(
            [str(python), "-c", "import venv, ensurepip"],
            capture_output=True, timeout=30, creationflags=NO_WINDOW,
        )
        return probe.returncode == 0
    except Exception:
        return False


def find_python() -> Path | None:
    """A system Python that can actually build — not this exe's own frozen
    runtime, and not a portable interpreter that only knows how to run."""
    candidates: list[Path] = []

    launcher = shutil.which("py")
    if launcher:
        try:
            out = subprocess.run(
                [launcher, "-3", "-c", "import sys; print(sys.executable)"],
                capture_output=True, text=True, timeout=20, creationflags=NO_WINDOW,
            )
            if out.returncode == 0 and out.stdout.strip():
                candidates.append(Path(out.stdout.strip()))
        except Exception:
            pass

    for name in ("python", "python3"):
        found = shutil.which(name)
        # The Microsoft Store stub on PATH is not an interpreter; it only opens
        # the Store page.
        if found and "WindowsApps" not in found:
            candidates.append(Path(found))

    for pattern in ("Python3*", "Python 3*"):
        for root in (Path(os.environ.get("LOCALAPPDATA", "")) / "Programs" / "Python",
                     Path("C:/"), Path("C:/Program Files")):
            if not root.exists():
                continue
            for directory in sorted(root.glob(pattern), reverse=True):
                exe = directory / "python.exe"
                if exe.exists():
                    candidates.append(exe)

    for candidate in candidates:
        if can_build(candidate):
            return candidate
    return None


ENTRY_NAMES = ("main.py", "__main__.py", "app.py", "run.py", "cli.py", "gui.py", "start.py")
SKIP_DIRS = {".git", ".venv", "venv", "env", "__pycache__", "build", "dist",
             "node_modules", ".idea", ".vscode", "tests", "test", "docs"}


def find_candidates(project: Path) -> list[Path]:
    """Everything in this repo that could plausibly be the thing to build,
    best guess first.

    Real repositories rarely hand you a single obvious file: the entry point
    might be src/app/__main__.py, or one of six scripts at the root. So this
    ranks what it finds and the window lets you override it — guessing wrong
    is fine, guessing and refusing to be corrected is not.
    """
    found: list[Path] = []

    def add(path: Path) -> None:
        if path.is_file() and path not in found:
            found.append(path)

    # A project's own spec beats anything this tool could reconstruct.
    for spec in sorted(project.glob("*.spec")):
        add(spec)
    for spec in sorted(project.glob("*/*.spec")):
        if spec.parent.name not in SKIP_DIRS:
            add(spec)

    # Conventional entry-point names, root first, then one and two levels in
    # (src/main.py, src/mypackage/__main__.py — the usual shapes).
    for name in ENTRY_NAMES:
        add(project / name)
    for depth in ("*", "*/*"):
        for name in ENTRY_NAMES:
            for path in sorted(project.glob(f"{depth}/{name}")):
                if not any(part in SKIP_DIRS for part in path.relative_to(project).parts):
                    add(path)

    # Anything else at the root, so a flat repo of scripts still offers a list.
    for path in sorted(project.glob("*.py")):
        if path.name not in ("__init__.py", "setup.py", "conftest.py"):
            add(path)

    return found


def describe(project: Path, entry: Path) -> str:
    relative = entry.relative_to(project) if entry.is_relative_to(project) else entry
    return f"{relative}  (spec)" if entry.suffix == ".spec" else str(relative)


def find_requirements(entry: Path, project: Path) -> Path | None:
    directory = entry.parent
    while True:
        candidate = directory / "requirements.txt"
        if candidate.exists():
            return candidate
        if directory == project or directory == directory.parent:
            return None
        directory = directory.parent


def read_build_args(entry: Path) -> list[str]:
    args_file = entry.parent / "build.args"
    if not args_file.exists():
        return []
    flags: list[str] = []
    for line in args_file.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if line and not line.startswith("#"):
            flags.extend(_split_args(line))
    return flags


def _split_args(line: str) -> list[str]:
    """--add-data "src;dest" has to survive as one argument."""
    out, current, quoted = [], "", False
    for char in line:
        if char == '"':
            quoted = not quoted
        elif char == " " and not quoted:
            if current:
                out.append(current)
                current = ""
        else:
            current += char
    if current:
        out.append(current)
    return out


# ---------------------------------------------------------------------------
# Building
# ---------------------------------------------------------------------------


class Builder:
    """Runs the build on a worker thread and reports progress as lines of text."""

    def __init__(self, log):
        self.log = log

    def run(self, cmd: list[str], cwd: Path | None = None) -> int:
        self.log("> " + " ".join(str(c) for c in cmd))
        process = subprocess.Popen(
            [str(c) for c in cmd], cwd=str(cwd) if cwd else None,
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, bufsize=1, errors="replace", creationflags=NO_WINDOW,
        )
        for line in process.stdout:
            self.log(line.rstrip())
        return process.wait()

    def ensure_venv(self, python: Path) -> Path:
        venv_python = VENV / "Scripts" / "python.exe"
        if not venv_python.exists():
            self.log(f"Creating build environment in {VENV}")
            HOME.mkdir(parents=True, exist_ok=True)
            if self.run([python, "-m", "venv", str(VENV)]) != 0:
                raise RuntimeError("Could not create the build environment")
        self.run([venv_python, "-m", "pip", "install", "--upgrade", "pip", "--quiet"])
        if self.run([venv_python, "-m", "pip", "install", PYINSTALLER_PIN, "--quiet"]) != 0:
            raise RuntimeError("Could not install PyInstaller")
        return venv_python

    def install_requirements(self, venv_python: Path, entry: Path, project: Path) -> None:
        reqs = find_requirements(entry, project)
        if reqs:
            self.log(f"Requirements: {reqs}")
            if self.run([venv_python, "-m", "pip", "install", "-r", str(reqs)]) != 0:
                raise RuntimeError(f"Some dependency in {reqs.name} would not install")
            return

        self.log("No requirements.txt — reading the imports instead (pipreqs)")
        self.run([venv_python, "-m", "pip", "install", "pipreqs", "--quiet"])
        generated = HOME / "generated-requirements.txt"
        self.run([venv_python, "-m", "pipreqs.pipreqs", str(entry.parent),
                  "--savepath", str(generated), "--force", "--mode", "no-pin"])
        if not generated.exists():
            return
        # One at a time: pipreqs infers package names from import names and gets
        # some wrong (win32com is pywin32, cv2 is opencv-python, a local module
        # looks like a package). One bad guess shouldn't stop the rest.
        for line in generated.read_text(encoding="utf-8").splitlines():
            package = line.strip()
            if package and not package.startswith("#"):
                if self.run([venv_python, "-m", "pip", "install", package]) != 0:
                    self.log(f"NOTE: could not install '{package}' — skipped. If the "
                             f"build fails on it, add a requirements.txt to the project.")

    def build(self, project: Path, entry: Path, mode: str, console: str, icon: str) -> Path:
        python = find_python()
        if python is None:
            raise RuntimeError(
                "No Python capable of building was found on this machine.\n\n"
                "PyInstaller needs a full interpreter to trace a program's imports. A "
                "portable or embeddable Python does not qualify — venv and ensurepip "
                "are stripped out of those, so they can run programs but not build "
                "them.\n\n"
                "Install Python from python.org, tick 'Add python.exe to PATH', and "
                "reopen this window. Or push the project and let the GitHub workflow "
                "build it, which needs nothing installed here."
            )
        self.log(f"Python: {python}")

        venv_python = self.ensure_venv(python)
        self.install_requirements(venv_python, entry, project)

        dist = project / "dist"
        WORKPATH.mkdir(parents=True, exist_ok=True)  # --specpath won't create it
        cmd = [venv_python, "-m", "PyInstaller", str(entry), "--noconfirm", "--clean",
               "--distpath", str(dist), "--workpath", str(WORKPATH)]

        if entry.suffix == ".spec":
            # Paths inside a spec resolve against the working directory, not the
            # spec's own, so a spec written to be run from its project folder
            # fails from anywhere else. distpath/workpath stay absolute.
            cwd = entry.parent
        else:
            cwd = None
            # A project folder's main.py should produce myproject.exe, not main.exe.
            name = entry.parent.name if entry.stem == "main" else entry.stem
            cmd += [f"--{mode}", "--paths", str(entry.parent),
                    "--specpath", str(WORKPATH), "--name", name]
            if console == "windowed":
                cmd.append("--windowed")
            if icon:
                cmd += ["--icon", icon]

        extra = read_build_args(entry)
        if extra:
            self.log(f"build.args: {' '.join(extra)}")
            cmd += extra

        if self.run(cmd, cwd=cwd) != 0:
            raise RuntimeError("PyInstaller failed — the reason is in the log above")

        built = sorted(dist.rglob("*.exe"), key=lambda p: p.stat().st_size, reverse=True)
        if not built:
            raise RuntimeError(f"Build reported success but produced no .exe in {dist}")
        return built[0]


# ---------------------------------------------------------------------------
# Window
# ---------------------------------------------------------------------------


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("860x600")
        self.minsize(700, 480)

        self.messages: queue.Queue[str] = queue.Queue()
        self.candidates: dict[str, Path] = {}
        self.entry_path: Path | None = None
        self.result: Path | None = None

        pad = {"padx": 8, "pady": 4}
        top = ttk.Frame(self)
        top.pack(fill="x", **pad)

        ttk.Label(top, text="Repo / project folder").grid(row=0, column=0, sticky="w")
        self.folder = tk.StringVar()
        ttk.Entry(top, textvariable=self.folder).grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(top, text="Browse...", command=self.pick_folder).grid(row=0, column=2)
        top.columnconfigure(1, weight=1)

        ttk.Label(top, text="Entry point").grid(row=1, column=0, sticky="w", pady=(6, 0))
        self.entry_choice = tk.StringVar()
        self.entry_box = ttk.Combobox(top, textvariable=self.entry_choice, state="readonly")
        self.entry_box.grid(row=1, column=1, sticky="ew", padx=6, pady=(6, 0))
        self.entry_box.bind("<<ComboboxSelected>>", self.on_entry_selected)
        ttk.Button(top, text="Pick file...",
                   command=self.pick_entry).grid(row=1, column=2, pady=(6, 0))

        self.detected = ttk.Label(top, text="Pick the folder holding your project.",
                                  foreground="#888")
        self.detected.grid(row=2, column=0, columnspan=3, sticky="w", pady=(6, 0))

        options = ttk.LabelFrame(self, text="Options")
        options.pack(fill="x", **pad)

        self.mode = tk.StringVar(value="onefile")
        self.console = tk.StringVar(value="console")
        self.icon = tk.StringVar()

        ttk.Radiobutton(options, text="One .exe", variable=self.mode,
                        value="onefile").grid(row=0, column=0, sticky="w", padx=8, pady=4)
        ttk.Radiobutton(options, text="Folder (starts faster)", variable=self.mode,
                        value="onedir").grid(row=0, column=1, sticky="w", padx=8)
        ttk.Radiobutton(options, text="Console", variable=self.console,
                        value="console").grid(row=0, column=2, sticky="w", padx=16)
        ttk.Radiobutton(options, text="Windowed (GUI)", variable=self.console,
                        value="windowed").grid(row=0, column=3, sticky="w", padx=8)

        ttk.Label(options, text="Icon (.ico)").grid(row=1, column=0, sticky="w", padx=8)
        ttk.Entry(options, textvariable=self.icon).grid(row=1, column=1, columnspan=2,
                                                        sticky="ew", padx=4, pady=4)
        ttk.Button(options, text="...", width=4,
                   command=self.pick_icon).grid(row=1, column=3, sticky="w")
        options.columnconfigure(1, weight=1)

        buttons = ttk.Frame(self)
        buttons.pack(fill="x", **pad)
        self.build_button = ttk.Button(buttons, text="Build .exe", command=self.start_build)
        self.build_button.pack(side="left")
        self.open_button = ttk.Button(buttons, text="Open output folder",
                                      command=self.open_output, state="disabled")
        self.open_button.pack(side="left", padx=6)
        self.status = ttk.Label(buttons, text="")
        self.status.pack(side="left", padx=10)

        log_frame = ttk.Frame(self)
        log_frame.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        self.log_box = tk.Text(log_frame, wrap="none", height=20, background="#111",
                               foreground="#ddd", insertbackground="#ddd")
        scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_box.yview)
        self.log_box.config(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.log_box.pack(side="left", fill="both", expand=True)

        self.after(80, self.drain)

    # --- plumbing ---------------------------------------------------------

    def log(self, message: str) -> None:
        self.messages.put(message)

    def drain(self) -> None:
        """The worker thread only ever puts strings on a queue; every widget
        touch happens here, on the UI thread, where tkinter requires it. The
        finished/failed signals travel the same queue as the log lines so they
        can't overtake the output they belong to."""
        while True:
            try:
                message = self.messages.get_nowait()
            except queue.Empty:
                break
            if message == DONE:
                self.finish(True)
            elif message == FAILED:
                self.finish(False)
            else:
                self.log_box.insert("end", message + "\n")
                self.log_box.see("end")
        self.after(80, self.drain)

    # --- actions ----------------------------------------------------------

    def pick_folder(self) -> None:
        chosen = filedialog.askdirectory(title="Select the repo or project folder")
        if chosen:
            self.folder.set(chosen)
            self.refresh_detection()

    def pick_entry(self) -> None:
        """The escape hatch for a layout the scan didn't anticipate."""
        project = Path(self.folder.get())
        chosen = filedialog.askopenfilename(
            title="Select the entry point",
            initialdir=str(project) if project.is_dir() else None,
            filetypes=[("Python or spec", "*.py *.spec"), ("All files", "*.*")],
        )
        if not chosen:
            return
        entry = Path(chosen)
        if not project.is_dir():
            # Choosing a file before a folder is a reasonable order to work in;
            # treat its folder as the project rather than complaining.
            project = entry.parent
            self.folder.set(str(project))
        label = describe(project, entry)
        self.candidates[label] = entry
        self.entry_box["values"] = list(self.candidates)
        self.entry_choice.set(label)
        self.on_entry_selected()

    def pick_icon(self) -> None:
        chosen = filedialog.askopenfilename(title="Select an icon",
                                            filetypes=[("Icon", "*.ico"), ("All files", "*.*")])
        if chosen:
            self.icon.set(chosen)

    def refresh_detection(self) -> None:
        project = Path(self.folder.get())
        if not project.is_dir():
            self.detected.config(text="That folder doesn't exist.", foreground="#c33")
            self.entry_path = None
            self.candidates = {}
            self.entry_box["values"] = []
            self.entry_choice.set("")
            return

        found = find_candidates(project)
        self.candidates = {describe(project, path): path for path in found}
        self.entry_box["values"] = list(self.candidates)

        if found:
            self.entry_choice.set(describe(project, found[0]))
            self.on_entry_selected()
        else:
            self.entry_path = None
            self.entry_choice.set("")
            self.detected.config(
                text="No .py or .spec found here — use 'Pick file...' to choose one.",
                foreground="#c33")

    def on_entry_selected(self, _event=None) -> None:
        self.entry_path = self.candidates.get(self.entry_choice.get())
        if self.entry_path is None:
            return
        if self.entry_path.suffix == ".spec":
            note = "the project's own spec decides mode, console and data files"
        else:
            note = f"exe will be named after {self.entry_path.parent.name}" \
                if self.entry_path.stem == "main" else "built with the options below"
        self.detected.config(text=f"Will build: {self.entry_path.name} — {note}",
                             foreground="#2a7")

    def start_build(self) -> None:
        # Deliberately not re-running detection here: that would discard an
        # entry point picked by hand and silently build something else.
        if self.entry_path is None:
            self.detected.config(text="Choose an entry point first.", foreground="#c33")
            return
        self.build_button.config(state="disabled")
        self.open_button.config(state="disabled")
        self.status.config(text="Building...", foreground="#888")
        self.log_box.delete("1.0", "end")
        threading.Thread(target=self.worker, daemon=True).start()

    def worker(self) -> None:
        builder = Builder(self.log)
        try:
            exe = builder.build(
                Path(self.folder.get()), self.entry_path,
                self.mode.get(), self.console.get(), self.icon.get().strip(),
            )
            size = exe.stat().st_size / 1024 / 1024
            self.result = exe
            self.log(f"\nBuilt {exe} ({size:.1f} MB)")
            self.messages.put(DONE)
        except Exception as error:  # shown in the window, never a silent stack trace
            self.log(f"\nFAILED: {error}")
            self.messages.put(FAILED)

    def finish(self, ok: bool) -> None:
        self.build_button.config(state="normal")
        if ok:
            self.open_button.config(state="normal")
            self.status.config(text="Done", foreground="#2a7")
        else:
            self.status.config(text="Failed — see the log", foreground="#c33")

    def open_output(self) -> None:
        if self.result:
            os.startfile(self.result.parent)  # noqa: S606 — Explorer on a known path


def main() -> int:
    App().mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
