"""
Builds the portable "WIR Generate Tools" folder.

    python tools\\deploy.py

The result is a folder you copy next to i125-WIR Cover Generate.xlsm on
any machine. Nothing is installed, nothing is compiled, and there is no
.bat, .vbs or .exe of ours for IT to object to - only Python's own
python.exe, which is signed by the Python Software Foundation.

    WIR Generate Tools\\
        WIR Generate Tools.lnk         double-click this
        Check this PC.lnk     run once on a new machine to prove it works
        runtime\\              a copy of the embeddable Python
        app\\                  main.py, core, ui
        README.txt

A .lnk can only hold an absolute path - there is no portable form of one.
So the folder needs one agreed location, given as --final-path, and the
shortcuts are written for it:

    a folder on the share, if the drive letter is the same for everyone,
    which also means one copy to update rather than many;

    or the same local path on every PC, such as C:\\WIR Generate Tools.

Built in place when --final-path is not given.
"""

from __future__ import annotations

import argparse
import os
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
PROJECT = os.path.dirname(HERE)

DEFAULT_RUNTIME = r"D:\i125\Document Management\python\python-3.13.15-embed-amd64"

# Copied into app\. Everything the program needs and nothing else.
APP_PARTS = ["main.py", "core", "ui"]

# Left out of runtime\. Development and documentation weight that a user
# machine never reads.
RUNTIME_SKIP = {
    "wheels",            # the pip wheels the runtime was built from
    "pythonwin",         # pywin32's IDE
    "PyWin32.chm",       # pywin32's help file
    "__pycache__",
}

README = """WIR Generate Tools
==================

Nothing here needs installing.

    1. Keep this folder where it was put. The shortcuts point at that
       exact path, so moving or renaming it stops them working - copy the
       shortcut to your desktop instead of moving the folder.
    2. Open i125-WIR Cover Generate.xlsm in Excel and select your WIR log
       sheet.
    3. Double-click "WIR Generate Tools".

On a machine that has not run this before, double-click "Check this PC"
first. It takes a few seconds, needs no workbook, and prints whether this
machine can run the program. It also writes app\\selftest.log, which is
the thing to send on if it reports a problem.

The runtime folder is a copy of Python, published by the Python Software
Foundation. The app folder is the program itself, in readable source form.
There is nothing compiled here.

The full manual is "WIR Generate Tools - User Manual.docx".
"""


def _copy_tree(source: str, target: str, skip: set[str] | None = None) -> int:
    """Copies a folder and reports how many files it wrote. shutil's own
    ignore callback is avoided so the skip list can be matched case-
    insensitively, which is what a Windows user would expect."""
    skip = {name.lower() for name in (skip or set())}
    written = 0

    for folder, subfolders, files in os.walk(source):
        subfolders[:] = [name for name in subfolders if name.lower() not in skip]

        relative = os.path.relpath(folder, source)
        destination = target if relative == "." else os.path.join(target, relative)
        os.makedirs(destination, exist_ok=True)

        for name in files:
            if name.lower() in skip or name.endswith(".pyc"):
                continue
            shutil.copy2(os.path.join(folder, name), os.path.join(destination, name))
            written += 1

    return written


def _make_shortcut(path: str, target: str, arguments: str, description: str,
                   root: str) -> None:
    """Writes a .lnk pointing at where the folder will finally live.

    A .lnk always stores an absolute target - there is no portable form of
    one. Windows will sometimes re-find a moved target on its own, but
    that is a heuristic and not something to hand to other people's
    machines, so the final location is stated up front instead (see
    --final-path) and the shortcut is simply correct for it.

    The working directory is deliberately left empty. Explorer then starts
    the process in pythonw.exe's own folder, which is what makes the
    script argument `..\\app\\main.py` resolve - so only the runtime path
    has to be right, not two paths.
    """
    import win32com.client

    shell = win32com.client.Dispatch("WScript.Shell")
    link = shell.CreateShortCut(path)
    link.TargetPath = target
    link.Arguments = arguments
    link.Description = description
    link.WorkingDirectory = ""
    link.Save()


def build(runtime_source: str, output: str, final: str) -> int:
    if not os.path.isdir(runtime_source):
        print(f"Runtime not found: {runtime_source}")
        print("Point --runtime at the embeddable Python folder.")
        return 1

    if not os.path.isfile(os.path.join(runtime_source, "pythonw.exe")):
        print(f"No pythonw.exe in {runtime_source} - that is not an embeddable Python.")
        return 1

    print(f"runtime from : {runtime_source}")
    print(f"app from     : {PROJECT}")
    print(f"building     : {output}")
    print(f"will live at : {final}")
    if os.path.normcase(output) != os.path.normcase(final):
        print("               (shortcuts are written for that path, not this one)")
    print()

    if os.path.isdir(output):
        shutil.rmtree(output)
    os.makedirs(output)

    runtime = os.path.join(output, "runtime")
    count = _copy_tree(runtime_source, runtime, RUNTIME_SKIP)
    print(f"  runtime : {count} files")

    app = os.path.join(output, "app")
    os.makedirs(app)
    count = 0
    for part in APP_PARTS:
        source = os.path.join(PROJECT, part)
        if os.path.isdir(source):
            count += _copy_tree(source, os.path.join(app, part), {"__pycache__"})
        elif os.path.isfile(source):
            shutil.copy2(source, os.path.join(app, part))
            count += 1
        else:
            print(f"  MISSING : {part}")
            return 1
    print(f"  app     : {count} files")

    manual = os.path.join(PROJECT, "WIR Generate Tools - User Manual.docx")
    if os.path.isfile(manual):
        shutil.copy2(manual, os.path.join(output, os.path.basename(manual)))
        print("  manual  : copied")

    with open(os.path.join(output, "README.txt"), "w", encoding="utf-8") as handle:
        handle.write(README)

    # Written for the final location, which is usually not where we are
    # building - so these two paths come from `final`, not `output`.
    final_runtime = os.path.join(final, "runtime")

    # pythonw for the window - python.exe would flash a console behind it.
    _make_shortcut(
        os.path.join(output, "WIR Generate Tools.lnk"),
        os.path.join(final_runtime, "pythonw.exe"),
        r'"..\app\main.py"',
        "Open WIR Generate Tools",
        final,
    )
    # python.exe for the check, so its output is visible, and --pause so
    # the window stays up long enough to read.
    _make_shortcut(
        os.path.join(output, "Check this PC.lnk"),
        os.path.join(final_runtime, "python.exe"),
        r'"..\app\main.py" --selftest --pause',
        "Prove this machine can run WIR Generate Tools",
        final,
    )
    print("  shortcuts: 2")

    total = sum(
        os.path.getsize(os.path.join(folder, name))
        for folder, _, files in os.walk(output)
        for name in files
    )
    print()
    print(f"Done. {total / 1_000_000:.0f} MB")
    print(f"Copy this folder to {final} - the shortcuts expect it there.")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Build the portable WIR Generate Tools folder.",
        epilog="The shortcuts hold an absolute path, so --final-path must be "
               "where the folder ends up. Use one location for everyone: a "
               "folder on the share, or the same local path on each PC.",
    )
    parser.add_argument("--runtime", default=DEFAULT_RUNTIME,
                        help="the embeddable Python to copy")
    parser.add_argument("--output", default=os.path.join(PROJECT, "dist", "WIR Generate Tools"),
                        help="where to build it")
    parser.add_argument("--final-path", default=None,
                        help="where the folder will live once copied "
                             "(defaults to --output, i.e. built in place)")
    options = parser.parse_args()

    output = os.path.abspath(options.output)
    final = os.path.abspath(options.final_path) if options.final_path else output
    return build(os.path.abspath(options.runtime), output, final)


if __name__ == "__main__":
    sys.exit(main())
