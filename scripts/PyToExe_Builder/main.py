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


def find_python() -> Path | None:
    """A system Python, not this exe's frozen runtime."""
    launcher = shutil.which("py")
    if launcher:
        try:
            out = subprocess.run(
                [launcher, "-3", "-c", "import sys; print(sys.executable)"],
                capture_output=True, text=True, timeout=20, creationflags=NO_WINDOW,
            )
            if out.returncode == 0 and out.stdout.strip():
                return Path(out.stdout.strip())
        except Exception:
            pass

    for name in ("python", "python3"):
        found = shutil.which(name)
        # The Microsoft Store stub on PATH is not a working interpreter; it
        # only opens the Store page.
        if found and "WindowsApps" not in found:
            return Path(found)

    for pattern in ("Python3*", "Python 3*"):
        for root in (Path(os.environ.get("LOCALAPPDATA", "")) / "Programs" / "Python",
                     Path("C:/"), Path("C:/Program Files")):
            if not root.exists():
                continue
            for candidate in sorted(root.glob(pattern), reverse=True):
                exe = candidate / "python.exe"
                if exe.exists():
                    return exe
    return None


def detect_entry(project: Path) -> tuple[Path | None, str]:
    """Same precedence as the workflow: a project's own .spec describes its
    build better than any flags this tool could pass."""
    specs = sorted(project.glob("*.spec"))
    if specs:
        return specs[0], f"spec: {specs[0].name}"

    main = project / "main.py"
    if main.exists():
        return main, "main.py"

    roots = [p for p in sorted(project.glob("*.py")) if p.name != "__init__.py"]
    if len(roots) == 1:
        return roots[0], roots[0].name
    if len(roots) > 1:
        return None, f"{len(roots)} .py files and no main.py — rename one to main.py"
    return None, "no .py or .spec found here"


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
                "No Python found on this machine.\n\n"
                "PyInstaller needs a real interpreter to trace a program's imports, so "
                "this tool cannot build without one.\n\n"
                "Install it from python.org and tick 'Add python.exe to PATH', then "
                "reopen this window."
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
        self.entry_path: Path | None = None
        self.result: Path | None = None

        pad = {"padx": 8, "pady": 4}
        top = ttk.Frame(self)
        top.pack(fill="x", **pad)

        ttk.Label(top, text="Project folder").grid(row=0, column=0, sticky="w")
        self.folder = tk.StringVar()
        ttk.Entry(top, textvariable=self.folder).grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(top, text="Browse...", command=self.pick_folder).grid(row=0, column=2)
        top.columnconfigure(1, weight=1)

        self.detected = ttk.Label(top, text="Pick the folder holding your project.",
                                  foreground="#888")
        self.detected.grid(row=1, column=0, columnspan=3, sticky="w", pady=(6, 0))

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
        chosen = filedialog.askdirectory(title="Select the project folder")
        if chosen:
            self.folder.set(chosen)
            self.refresh_detection()

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
            return
        entry, description = detect_entry(project)
        self.entry_path = entry
        if entry:
            self.detected.config(text=f"Will build: {description}", foreground="#2a7")
        else:
            self.detected.config(text=description, foreground="#c33")

    def start_build(self) -> None:
        self.refresh_detection()
        if self.entry_path is None:
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
